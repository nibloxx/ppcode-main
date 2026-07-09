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
    """Use Esri for geocoding, imagery, and demographics; Google for Street View."""

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
        if self.esri_api_key:
            try:
                return self._get_aerial_esri(lat, lng, output_path)
            except Exception as exc:
                logger.warning("Esri aerial image failed, trying Google fallback: %s", exc)

        if self.google_api_key:
            return self._get_aerial_google(lat, lng, output_path)

        logger.warning("No aerial imagery provider available")
        return None

    def get_street_view_image(self, address: str, output_path: Path) -> Optional[str]:
        if not self.google_api_key:
            logger.info("Skipping Street View (GOOGLE_API_KEY not configured)")
            return None

        params = {
            "size": "600x500",
            "location": address,
            "pitch": "0",
            "fov": "90",
            "key": self.google_api_key,
        }
        response = requests.get(GOOGLE_STREETVIEW_URL, params=params, timeout=30)
        response.raise_for_status()

        if response.headers.get("content-type", "").startswith("image/"):
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_bytes(response.content)
            return str(output_path)

        logger.warning("Street View not available for this address")
        return None

    def get_demographics(self, lat: float, lng: float) -> Dict:
        if not self.esri_api_key:
            return {}

        try:
            return self._get_demographics_esri(lat, lng)
        except Exception as exc:
            logger.warning("Esri demographics unavailable, caller may use AI fallback: %s", exc)
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
        params = {
            "center": f"{lat},{lng}",
            "zoom": "18",
            "size": "1920x1920",
            "maptype": "hybrid",
            "key": self.google_api_key,
            "markers": f"{lat},{lng}",
        }
        response = requests.get(GOOGLE_STATICMAP_URL, params=params, timeout=30)
        response.raise_for_status()

        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(response.content)
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
