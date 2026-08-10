"""Hybrid Esri + Google location services for property reports."""
import json
import logging
from pathlib import Path
from typing import Dict, Optional, Tuple

import requests

logger = logging.getLogger(__name__)

ESRI_GEOCODE_URL = (
    "https://geocode-api.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates"
)
ESRI_IMAGERY_EXPORT_URL = (
    "https://services.arcgisonline.com/arcgis/rest/services/World_Imagery/MapServer/export"
)
ESRI_ENRICH_URL = (
    "https://geoenrich.arcgis.com/arcgis/rest/services/World/GeoenrichmentServer/Geoenrichment/enrich"
)
GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"
GOOGLE_STREETVIEW_URL = "https://maps.googleapis.com/maps/api/streetview"
GOOGLE_STATICMAP_URL = "https://maps.googleapis.com/maps/api/staticmap"


class HybridLocationService:
    """Esri for geocoding/demographics; Google for aerial + Street View imagery."""

    def __init__(
        self,
        esri_api_key: Optional[str] = None,
        google_api_key: Optional[str] = None,
    ):
        self.esri_api_key = (esri_api_key or "").strip() or None
        self.google_api_key = (google_api_key or "").strip() or None

        if not self.esri_api_key and not self.google_api_key:
            raise ValueError(
                "At least one location API key is required (ESRI_API_KEY or GOOGLE_API_KEY)"
            )

    def geocode(self, address: str) -> Tuple[float, float, Dict]:
        esri_error = None

        if self.esri_api_key:
            try:
                return self._geocode_esri(address)
            except Exception as exc:
                esri_error = str(exc)
                logger.warning("Esri geocoding failed, trying Google fallback: %s", exc)

        if self.google_api_key:
            try:
                return self._geocode_google(address)
            except Exception as exc:
                google_error = str(exc)
                if esri_error:
                    raise ValueError(
                        f"Geocoding failed for '{address}'. Esri: {esri_error}. Google: {google_error}"
                    ) from exc
                raise

        if esri_error:
            raise ValueError(
                f"Geocoding failed for '{address}'. Esri: {esri_error}. "
                "Set a valid GOOGLE_API_KEY for fallback geocoding."
            )

        raise ValueError("No geocoding provider available")

    def get_aerial_image(self, lat: float, lng: float, output_path: Path) -> Optional[str]:
        """Prefer Google Static Maps (hybrid) for aerial; fall back to Esri."""
        if self.google_api_key:
            try:
                return self._get_aerial_google(lat, lng, output_path)
            except Exception as exc:
                logger.warning("Google aerial image failed, trying Esri fallback: %s", exc)

        if self.esri_api_key:
            try:
                return self._get_aerial_esri(lat, lng, output_path)
            except Exception as exc:
                logger.warning("Esri aerial image failed: %s", exc)

        logger.warning("No aerial imagery provider available")
        return None

    def get_street_view_image(
        self,
        address: str,
        output_path: Path,
        lat: Optional[float] = None,
        lng: Optional[float] = None,
    ) -> Optional[str]:
        """Fetch a real Google Street View photo (never an aerial / static map)."""
        if not self.google_api_key:
            logger.info("Skipping Street View (GOOGLE_API_KEY not configured)")
            return None

        # Prefer the typed address so imagery matches what the user entered.
        # Lat/lng is a fallback when address lookup has no Street View coverage.
        locations = []
        if address:
            locations.append(address.strip())
        if lat is not None and lng is not None:
            locations.append(f"{lat},{lng}")
        # de-dupe while preserving order
        seen = set()
        locations = [loc for loc in locations if not (loc in seen or seen.add(loc))]

        logger.info(
            "Street View lookup for address=%r (also trying coords if needed)",
            address,
        )

        last_error = None
        for location in locations:
            # Metadata: only hard-skip when Google explicitly says no imagery.
            # Timeouts/network errors should still try the image download.
            meta = self._street_view_metadata(location)
            if meta == "ZERO_RESULTS" or meta == "NOT_FOUND":
                logger.warning(
                    "Street View metadata=%s for %s — trying next location",
                    meta,
                    location,
                )
                continue

            for source in ("outdoor", None):
                try:
                    params = {
                        "size": "640x640",
                        "location": location,
                        "pitch": "0",
                        "fov": "90",
                        "key": self.google_api_key,
                    }
                    if source:
                        params["source"] = source
                    response = None
                    for attempt in range(1, 4):
                        try:
                            response = requests.get(
                                GOOGLE_STREETVIEW_URL, params=params, timeout=45
                            )
                            response.raise_for_status()
                            break
                        except Exception as exc:
                            last_error = exc
                            logger.warning(
                                "Street View fetch attempt %s failed for %s: %s",
                                attempt,
                                location,
                                exc,
                            )
                            if attempt < 3:
                                continue
                            response = None
                    if response is None:
                        continue

                    content_type = response.headers.get("content-type", "")
                    if not content_type.startswith("image/"):
                        logger.warning(
                            "Street View did not return an image for %s", location
                        )
                        continue

                    # Reject tiny placeholder responses
                    if len(response.content) < 8000:
                        logger.warning(
                            "Street View response looks like a placeholder for %s",
                            location,
                        )
                        continue

                    output_path = Path(output_path)
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    data = response.content
                    try:
                        from io import BytesIO
                        from PIL import Image

                        img = Image.open(BytesIO(data)).convert("RGB")
                        buf = BytesIO()
                        img.save(buf, format="JPEG", quality=75, optimize=True)
                        data = buf.getvalue()
                    except Exception as exc:
                        logger.warning("Could not recompress Street View image: %s", exc)
                    output_path.write_bytes(data)
                    logger.info(
                        "Street View saved (%s bytes) for %s source=%s",
                        len(data),
                        location,
                        source or "default",
                    )
                    return str(output_path)
                except Exception as exc:
                    last_error = exc
                    logger.warning("Street View attempt failed for %s: %s", location, exc)

        logger.warning(
            "No Street View image available after retries (last error: %s)", last_error
        )
        return None

    def _street_view_metadata(self, location: str) -> Optional[str]:
        """Return Google metadata status, or None if the check itself failed."""
        try:
            response = requests.get(
                f"{GOOGLE_STREETVIEW_URL}/metadata",
                params={
                    "location": location,
                    "source": "outdoor",
                    "key": self.google_api_key,
                },
                timeout=45,
            )
            response.raise_for_status()
            status = (response.json() or {}).get("status", "")
            return status or None
        except Exception as exc:
            logger.warning(
                "Street View metadata check failed (will still try image fetch): %s",
                exc,
            )
            return None

    def _street_view_available(self, location: str) -> bool:
        status = self._street_view_metadata(location)
        return status == "OK"

    def get_demographics(self, lat: float, lng: float) -> Dict:
        if not self.esri_api_key:
            return {}

        try:
            return self._get_demographics_esri(lat, lng)
        except Exception as exc:
            logger.warning("Esri demographics unavailable, caller may use AI fallback: %s", exc)
            return {}

    def get_admin_demographics(self, lat: float, lng: float) -> Dict[str, Dict]:
        """Esri GeoEnrichment for US / State / County intersecting the point.

        Returns {"us": {...}, "state": {...}, "county": {...}} with normalized
        keys used by BOV population / household / tenure tables. Empty {} on failure.
        """
        if not self.esri_api_key:
            return {}

        try:
            return self._get_admin_demographics_esri(lat, lng)
        except Exception as exc:
            logger.warning("Esri admin demographics unavailable: %s", exc)
            return {}

    def get_ring_demographics(
        self, lat: float, lng: float, radii=(1, 3, 5)
    ) -> Dict[str, Dict]:
        """Return Esri GeoEnrichment demographics for 1/3/5-mile rings.

        Result shape: {"1": {var: value, ...}, "3": {...}, "5": {...}}.
        Returns {} when Esri is unavailable so callers can fall back to AI.
        """
        if not self.esri_api_key:
            return {}

        try:
            return self._get_ring_demographics_esri(lat, lng, list(radii))
        except Exception as exc:
            logger.warning("Esri ring demographics unavailable, using fallback: %s", exc)
            return {}

    def _get_ring_demographics_esri(self, lat: float, lng: float, radii) -> Dict[str, Dict]:
        study_areas = [{"geometry": {"x": lng, "y": lat, "spatialReference": {"wkid": 4326}}}]
        params = {
            "f": "json",
            "token": self.esri_api_key,
            "studyAreas": json.dumps(study_areas),
            "studyAreasOptions": json.dumps(
                {"areaType": "RingBuffer", "bufferUnits": "esriMiles", "bufferRadii": radii}
            ),
            "returnGeometry": "false",
            "dataCollections": json.dumps(["KeyUSFacts", "Age", "Income", "Housing"]),
        }
        response = requests.post(ESRI_ENRICH_URL, data=params, timeout=45)
        response.raise_for_status()
        data = response.json()
        self._check_esri_response(data, "ring enrichment")

        feature_sets = data.get("results", [{}])[0].get("value", {}).get("FeatureSet", [])
        rings: Dict[str, Dict] = {}
        for feature_set in feature_sets:
            for record in feature_set.get("features", []):
                attributes = record.get("attributes", {})
                radius = attributes.get("AREA_DESC") or attributes.get("bufferRadii")
                radius_key = self._ring_key(radius, len(rings), radii)
                rings.setdefault(radius_key, {}).update(attributes)
        return rings

    def _get_admin_demographics_esri(self, lat: float, lng: float) -> Dict[str, Dict]:
        """Pull KeyUSFacts for WholeUSA / State / County at the property point."""
        analysis_vars = [
            "KeyUSFacts.TOTPOP_CY",
            "KeyUSFacts.TOTPOP_FY",
            "KeyUSFacts.TOTPOP20",
            "KeyUSFacts.TOTPOP10",
            "KeyUSFacts.TOTHH_CY",
            "KeyUSFacts.TOTHH_FY",
            "KeyUSFacts.AVGHHSZ_CY",
            "KeyUSFacts.OWNER_CY",
            "KeyUSFacts.RENTER_CY",
            "KeyUSFacts.TOTHU_CY",
        ]
        params = {
            "f": "json",
            "token": self.esri_api_key,
            "returnGeometry": "false",
            "analysisVariables": json.dumps(analysis_vars),
            "studyAreas": json.dumps(
                [
                    {
                        "geometry": {
                            "x": lng,
                            "y": lat,
                            "spatialReference": {"wkid": 4326},
                        },
                        "comparisonLevels": [
                            {"layer": "US.WholeUSA"},
                            {"layer": "US.States"},
                            {"layer": "US.Counties"},
                        ],
                    }
                ]
            ),
        }
        response = requests.post(ESRI_ENRICH_URL, data=params, timeout=60)
        response.raise_for_status()
        data = response.json()
        self._check_esri_response(data, "admin enrichment")

        features = (
            data.get("results", [{}])[0]
            .get("value", {})
            .get("FeatureSet", [{}])[0]
            .get("features", [])
        )
        land_sq_mi = self._get_land_areas_sq_mi(lat, lng)
        out: Dict[str, Dict] = {}
        for record in features:
            attrs = record.get("attributes", {}) or {}
            level = (attrs.get("StdGeographyLevel") or "").upper()
            name = attrs.get("StdGeographyName") or ""
            if "WHOLEUSA" in level or name.lower() == "united states":
                geo_key = "us"
            elif "STATES" in level:
                geo_key = "state"
            elif "COUNTIES" in level:
                geo_key = "county"
            else:
                # Skip the local block-group / point feature (no StdGeography*)
                continue
            out[geo_key] = self._normalize_admin_attrs(attrs, land_sq_mi.get(geo_key))
        return out

    @staticmethod
    def _normalize_admin_attrs(attrs: Dict, land_sq_mi: Optional[float]) -> Dict:
        def _num(*names):
            for name in names:
                if name in attrs and attrs[name] is not None:
                    try:
                        return float(attrs[name])
                    except (TypeError, ValueError):
                        continue
            return None

        pop_2010 = _num("TOTPOP10", "KeyUSFacts.TOTPOP10")
        pop_2020 = _num("TOTPOP20", "KeyUSFacts.TOTPOP20")
        pop_cy = _num("TOTPOP_CY", "KeyUSFacts.TOTPOP_CY")
        pop_fy = _num("TOTPOP_FY", "KeyUSFacts.TOTPOP_FY")
        hh_cy = _num("TOTHH_CY", "KeyUSFacts.TOTHH_CY")
        hh_fy = _num("TOTHH_FY", "KeyUSFacts.TOTHH_FY")
        hhsize_cy = _num("AVGHHSZ_CY", "KeyUSFacts.AVGHHSZ_CY")
        owner = _num("OWNER_CY", "KeyUSFacts.OWNER_CY")
        renter = _num("RENTER_CY", "KeyUSFacts.RENTER_CY")
        hu_cy = _num("TOTHU_CY", "KeyUSFacts.TOTHU_CY")

        hhsize_fy = None
        if pop_fy and hh_fy and hh_fy > 0:
            hhsize_fy = round(pop_fy / hh_fy, 2)

        owner_pct = renter_pct = None
        if owner is not None and renter is not None and (owner + renter) > 0:
            total = owner + renter
            owner_pct = round(100.0 * owner / total, 1)
            renter_pct = round(100.0 * renter / total, 1)

        dens_2020 = dens_cy = None
        if land_sq_mi and land_sq_mi > 0:
            if pop_2020 is not None:
                dens_2020 = round(pop_2020 / land_sq_mi)
            if pop_cy is not None:
                dens_cy = round(pop_cy / land_sq_mi)

        return {
            "name": attrs.get("StdGeographyName"),
            "pop_2010": int(pop_2010) if pop_2010 is not None else None,
            "pop_2020": int(pop_2020) if pop_2020 is not None else None,
            "pop_cy": int(pop_cy) if pop_cy is not None else None,
            "pop_fy": int(pop_fy) if pop_fy is not None else None,
            "hh_cy": int(hh_cy) if hh_cy is not None else None,
            "hh_fy": int(hh_fy) if hh_fy is not None else None,
            "hhsize_cy": round(hhsize_cy, 2) if hhsize_cy is not None else None,
            "hhsize_fy": hhsize_fy,
            "owner_pct": owner_pct,
            "renter_pct": renter_pct,
            "housing_units_cy": int(hu_cy) if hu_cy is not None else None,
            "density_2020": dens_2020,
            "density_cy": dens_cy,
            "land_sq_mi": land_sq_mi,
        }

    def _get_land_areas_sq_mi(self, lat: float, lng: float) -> Dict[str, float]:
        """Census TIGERWeb land area (m² → sq mi) for state/county; US is fixed."""
        # U.S. Census Bureau figure for contiguous + AK/HI land area (sq mi)
        areas: Dict[str, float] = {"us": 3531905.43}
        sqm_to_sqmi = 2589988.110336

        tiger = (
            "https://tigerweb.geo.census.gov/arcgis/rest/services/"
            "TIGERweb/State_County/MapServer/{layer}/query"
        )
        point = {
            "geometry": f"{lng},{lat}",
            "geometryType": "esriGeometryPoint",
            "inSR": 4326,
            "spatialRel": "esriSpatialRelIntersects",
            "returnGeometry": "false",
            "f": "json",
        }
        try:
            state_resp = requests.get(
                tiger.format(layer=0),
                params={**point, "outFields": "NAME,AREALAND"},
                timeout=30,
            )
            state_resp.raise_for_status()
            state_feats = state_resp.json().get("features") or []
            if state_feats:
                land = float(state_feats[0]["attributes"]["AREALAND"])
                areas["state"] = land / sqm_to_sqmi

            county_resp = requests.get(
                tiger.format(layer=1),
                params={**point, "outFields": "NAME,AREALAND"},
                timeout=30,
            )
            county_resp.raise_for_status()
            county_feats = county_resp.json().get("features") or []
            if county_feats:
                land = float(county_feats[0]["attributes"]["AREALAND"])
                areas["county"] = land / sqm_to_sqmi
        except Exception as exc:
            logger.warning("Census TIGER land area lookup failed: %s", exc)

        return areas

    @staticmethod
    def _ring_key(radius, index, radii) -> str:
        if isinstance(radius, str):
            for candidate in radii:
                if str(candidate) in radius:
                    return str(candidate)
        if index < len(radii):
            return str(radii[index])
        return str(radius)

    def _check_esri_response(self, data: dict, context: str) -> None:
        if error := data.get("error"):
            code = error.get("code", "unknown")
            message = error.get("message", "Esri request failed")
            hint = ""
            if code == 498 or "invalid token" in message.lower():
                hint = (
                    " Use a permanent API key from https://developers.arcgis.com/ "
                    "(not an expired OAuth access token)."
                )
            raise ValueError(f"Esri {context} error ({code}): {message}.{hint}")

    def _geocode_esri(self, address: str) -> Tuple[float, float, Dict]:
        params = {
            "f": "json",
            "singleLine": address,
            "countryCode": "USA",
            "outFields": "City,Region,Subregion,Postal,Addr_type",
            "token": self.esri_api_key,
        }
        response = requests.get(ESRI_GEOCODE_URL, params=params, timeout=30)
        response.raise_for_status()
        data = response.json()
        self._check_esri_response(data, "geocoding")

        candidates = data.get("candidates") or []
        if not candidates:
            raise ValueError(f"No Esri geocoding results for address: {address}")

        best = candidates[0]
        location = best["location"]
        attributes = best.get("attributes", {})

        details = {
            "city": attributes.get("City", ""),
            "state": attributes.get("Region", ""),
            "county": attributes.get("Subregion", ""),
            "zip_code": attributes.get("Postal", ""),
        }
        return float(location["y"]), float(location["x"]), details

    def _geocode_google(self, address: str) -> Tuple[float, float, Dict]:
        params = {"address": address, "key": self.google_api_key}
        response = requests.get(GOOGLE_GEOCODE_URL, params=params, timeout=30)
        response.raise_for_status()
        data = response.json()

        if data.get("status") != "OK" or not data.get("results"):
            detail = data.get("error_message") or data.get("status")
            raise ValueError(
                f"Google geocoding failed for address: {address} ({detail})"
            )

        result = data["results"][0]
        location = result["geometry"]["location"]
        details = {}

        for component in result.get("address_components", []):
            types = component.get("types", [])
            if "administrative_area_level_2" in types:
                details["county"] = component["long_name"]
            elif "administrative_area_level_1" in types:
                details["state"] = component["long_name"]
            elif "locality" in types:
                details["city"] = component["long_name"]
            elif "postal_code" in types:
                details["zip_code"] = component["long_name"]

        return float(location["lat"]), float(location["lng"]), details

    def _get_aerial_esri(self, lat: float, lng: float, output_path: Path) -> str:
        delta = 0.0015
        params = {
            "bbox": f"{lng - delta},{lat - delta},{lng + delta},{lat + delta}",
            "bboxSR": "4326",
            "imageSR": "4326",
            "size": "1920,1920",
            "format": "jpg",
            "f": "image",
            "token": self.esri_api_key,
        }
        response = requests.get(ESRI_IMAGERY_EXPORT_URL, params=params, timeout=60)
        response.raise_for_status()

        if not response.headers.get("content-type", "").startswith("image/"):
            raise ValueError("Esri imagery export did not return an image")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(response.content)
        return str(output_path)

    def _get_aerial_google(self, lat: float, lng: float, output_path: Path) -> str:
        """Google Static Maps hybrid aerial (satellite + labels + pin), saved as JPEG."""
        from io import BytesIO

        params = {
            "center": f"{lat},{lng}",
            "zoom": "18",
            "size": "640x640",
            "scale": "1",
            "maptype": "hybrid",
            "key": self.google_api_key,
            "markers": f"color:red|{lat},{lng}",
        }
        response = requests.get(GOOGLE_STATICMAP_URL, params=params, timeout=30)
        response.raise_for_status()

        if not response.headers.get("content-type", "").startswith("image/"):
            raise ValueError("Google Static Map did not return an image")

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        data = response.content
        # Static Maps returns PNG; Word template slots expect JPEG
        if data[:4] == b"\x89PNG":
            try:
                from PIL import Image

                img = Image.open(BytesIO(data)).convert("RGB")
                buf = BytesIO()
                img.save(buf, format="JPEG", quality=75, optimize=True)
                data = buf.getvalue()
            except Exception as exc:
                logger.warning("Could not convert Google aerial PNG to JPEG: %s", exc)

        output_path.write_bytes(data)
        return str(output_path)

    def _get_demographics_esri(self, lat: float, lng: float) -> Dict:
        study_areas = [
            {
                "geometry": {
                    "x": lng,
                    "y": lat,
                    "spatialReference": {"wkid": 4326},
                }
            }
        ]
        params = {
            "f": "json",
            "token": self.esri_api_key,
            "studyAreas": json.dumps(study_areas),
            "returnGeometry": "false",
            "analysisVariables": json.dumps(
                [
                    "KeyUSFacts.TOTPOP_CY",
                    "KeyUSFacts.POPGRWCYFY",
                    "KeyUSFacts.TOTHH_CY",
                    "KeyUSFacts.AVGHHSZ_CY",
                    "KeyUSFacts.EMP_CY",
                    "KeyUSFacts.UNEMPRT_CY",
                    "KeyUSFacts.MEDHINC_CY",
                ]
            ),
        }
        response = requests.post(ESRI_ENRICH_URL, data=params, timeout=30)
        response.raise_for_status()
        data = response.json()

        if data.get("error"):
            raise ValueError(data["error"].get("message", "Esri enrichment failed"))

        features = data.get("results", [{}])[0].get("value", {}).get("FeatureSet", [{}])
        records = features[0].get("features", [])
        if not records:
            return {}

        attributes = records[0].get("attributes", {})

        def _attr(*names):
            for name in names:
                if name in attributes and attributes[name] is not None:
                    return attributes[name]
            return None

        total_pop = _attr("KeyUSFacts.TOTPOP_CY", "TOTPOP_CY")
        growth = _attr("KeyUSFacts.POPGRWCYFY", "POPGRWCYFY")
        households = _attr("KeyUSFacts.TOTHH_CY", "TOTHH_CY")
        avg_household = _attr("KeyUSFacts.AVGHHSZ_CY", "AVGHHSZ_CY")
        employment = _attr("KeyUSFacts.EMP_CY", "EMP_CY")
        unemployment = _attr("KeyUSFacts.UNEMPRT_CY", "UNEMPRT_CY")
        median_income = _attr("KeyUSFacts.MEDHINC_CY", "MEDHINC_CY")

        demographics = {
            "population_current": int(total_pop) if total_pop is not None else None,
            "population_2020": int(total_pop) if total_pop is not None else None,
            "population_growth_rate": round(float(growth), 2) if growth is not None else None,
            "households_2020": int(households) if households is not None else None,
            "avg_household_size": round(float(avg_household), 2) if avg_household is not None else None,
            "employment_count": int(employment) if employment is not None else None,
            "unemployment_rate": round(float(unemployment), 2) if unemployment is not None else None,
            "major_industries": [],
            "median_income": int(median_income) if median_income is not None else None,
            "data_source": "Esri GeoEnrichment",
        }

        if total_pop and employment and demographics.get("unemployment_rate") is None:
            demographics["employment_rate"] = round((float(employment) / float(total_pop)) * 100, 1)

        return {key: value for key, value in demographics.items() if value is not None}
