import openai
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from datetime import datetime
import os
import json
import requests
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, field
import logging
from pathlib import Path
import re
from location_services import HybridLocationService

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

@dataclass
class PropertyReportData:
    """Complete data structure for property report"""
    # Basic Info
    address: str
    date: str = field(default_factory=lambda: datetime.now().strftime("%B %d, %Y"))
    
    # Preparer Info
    prepared_by: str = ""
    prepared_by_title: str = ""
    prepared_by_company: str = ""
    prepared_by_address: str = ""
    
    # Client Info  
    prepared_for: str = ""
    prepared_for_title: str = ""
    prepared_for_company: str = ""
    prepared_for_address: str = ""
    
    # Property Details
    property_name: str = ""
    property_type: str = ""
    state: str = ""
    county: str = ""
    longitude: str = ""
    latitude: str = ""
    
    # Physical Characteristics (mostly static)
    topography: str = "Level at street Grade"
    shape: str = "Irregular"
    access: str = "Good"  # Average, Average/Good, Good, Good/Excellent, Excellent
    exposure: str = "Average/Good"  # Average, Average/Good, Good, Good/Excellent, Excellent
    
    # Property Specific
    lot_area: str = ""
    acres: str = ""
    recorded_sale_date: str = ""
    zoning: str = ""
    apn: str = ""
    current_owner: str = ""
    
    # Transaction context
    lease_or_sale: str = ""

    # Market Analysis (static)
    marketing_period: str = "Six months or less"
    
    # SWOT Analysis
    swot_strengths: str = ""
    swot_weaknesses: str = ""
    swot_opportunities: str = ""
    swot_threats: str = ""
    
    # Generated Content
    property_summary: str = ""
    location_summary: str = ""
    demographic_analysis: str = ""
    size_and_topography: str = ""
    population_analysis: str = ""
    household_trends: str = ""
    employment_analysis: str = ""
    economic_factors: str = ""
    community_services: str = ""
    
    # Market Analysis fields
    market_overview: str = ""
    vacancy_rates: str = ""
    lease_rates: str = ""
    construction_activity: str = ""
    market_trends: str = ""
    investment_insights: str = ""
    market_recommendations: str = ""
    market_data_sources: str = ""
    market_quarter: str = field(default_factory=lambda: f"Q{(datetime.now().month-1)//3 + 1} {datetime.now().year}")
    
    # Image paths
    aerial_image_path: Optional[str] = None
    street_view_image_path: Optional[str] = None

    # BOV table values (population, households, rings, employment, valuation, etc.)
    table_values: Dict[str, str] = field(default_factory=dict)

    # Additional BOV narrative sections
    executive_summary: str = ""
    regional_analysis: str = ""
    sales_conclusion: str = ""
    reconciliation_summary: str = ""
    reconciliation_notes: str = ""

    # Comparable sales extracted from uploaded CoStar PDF (optional)
    comps: List[Any] = field(default_factory=list)
    # Raw demographic dataset used for employment table refresh
    bov_dataset: Dict = field(default_factory=dict)

class ComprehensivePropertyReportGenerator:
    def __init__(
        self,
        openai_api_key: str,
        template_path: str,
        output_dir: str = "output",
        google_api_key: Optional[str] = None,
        esri_api_key: Optional[str] = None,
    ):
        """
        Initialize the Comprehensive Property Report Generator
        
        Args:
            openai_api_key: OpenAI API key
            template_path: Path to the Word document template
            output_dir: Directory to save generated reports
            google_api_key: Google Maps API key (Street View; optional fallback)
            esri_api_key: Esri API key (geocoding, aerial imagery, demographics)
        """
        self.openai_client = openai.OpenAI(api_key=openai_api_key)
        self.location_service = HybridLocationService(
            esri_api_key=esri_api_key,
            google_api_key=google_api_key,
        )
        self.template_path = Path(template_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        # Create images subdirectory
        self.images_dir = self.output_dir / "images"
        self.images_dir.mkdir(exist_ok=True)
        
        # Verify template exists
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template file not found: {template_path}")

    def get_property_images(self, address: str, lat: float, lng: float) -> Tuple[Optional[str], Optional[str]]:
        """
        Get aerial and street view images for the property
        
        Args:
            address: The property address
            lat: Latitude
            lng: Longitude
            
        Returns:
            Tuple of (aerial_image_path, street_view_image_path)
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        aerial_path = None
        street_view_path = None
        
        try:
            logger.info("Fetching aerial image via Google Static Maps (Esri fallback)...")
            aerial_filename = f"aerial_{timestamp}.jpg"
            aerial_path = self.location_service.get_aerial_image(
                lat, lng, self.images_dir / aerial_filename
            )
            if aerial_path:
                logger.info("Aerial image saved: %s", aerial_path)

            logger.info("Fetching Street View image via Google...")
            street_view_filename = f"street_view_{timestamp}.jpg"
            street_view_path = self.location_service.get_street_view_image(
                address,
                self.images_dir / street_view_filename,
                lat=lat,
                lng=lng,
            )
            if street_view_path:
                logger.info("Street view image saved: %s", street_view_path)
            else:
                logger.warning("No Street View image available; SUBJECT PHOTOS will be empty")
                
        except Exception as e:
            logger.error(f"Error fetching images: {e}")
            
        return aerial_path, street_view_path

    def get_coordinates_and_details(self, address: str) -> Tuple[float, float, Dict]:
        """
        Get latitude, longitude, and location details (Esri first, Google fallback).
        """
        try:
            return self.location_service.geocode(address)
        except Exception as e:
            logger.error(f"Error getting coordinates for {address}: {e}")
            raise

    def get_census_data(self, lat: float, lng: float, county: str, state: str) -> Dict:
        """
        Get demographic data from Esri GeoEnrichment, with AI fallback.
        """
        try:
            demographics = self.location_service.get_demographics(lat, lng)
            if demographics:
                logger.info("Using Esri GeoEnrichment demographics")
                return demographics

            census_prompt = f"""
            Generate realistic and current demographic data for coordinates {lat}, {lng} in {county}, {state}.
            Include population statistics, household data, employment data, and economic factors.
            Format as JSON with keys: population_2020, population_growth_rate, households_2020, 
            avg_household_size, employment_rate, major_industries, median_income.
            """
            
            response = self.openai_client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a demographic data analyst. Provide realistic census-like data in JSON format."},
                    {"role": "user", "content": census_prompt}
                ],
                temperature=0.1
            )
            
            return json.loads(response.choices[0].message.content)
            
        except Exception as e:
            logger.error(f"Error getting census data: {e}")
            return {}

    def _search_current_market_data(self, county: str, state: str, property_type: str) -> Dict:
        """Search for current market data using AI to simulate real market conditions"""
        
        # Get current quarter
        current_date = datetime.now()
        current_quarter = f"Q{(current_date.month-1)//3 + 1} {current_date.year}"
        
        prompt = f"""
        Generate realistic current {current_quarter} market data for {property_type} in {county}, {state}.
        
        Include the following metrics with realistic values:
        1. Direct vacancy rate with QoQ and YoY changes
        2. Sublease vacancy rate with QoQ change
        3. Total vacancy rate
        4. Average lease rates overall and by class (A, B, C)
        5. Construction pipeline (square footage under construction and type)
        6. Market absorption rates
        7. Population growth projections
        8. Employment statistics
        9. Major market trends
        
        Format as JSON with this structure:
        {{
            "quarter": "{current_quarter}",
            "direct_vacancy": 12.71,
            "direct_qoq": "+1.02",
            "direct_yoy": "+1.68",
            "sublease_vacancy": 5.51,
            "sublease_qoq": "-0.39",
            "total_vacancy": 18.22,
            "avg_lease_rate": 24.33,
            "lease_rate_yoy": "-0.04",
            "cap_rate": 6.7,
            "class_a_rate": 27.13,
            "class_b_rate": 22.03,
            "class_c_rate": 19.44,
            "construction_sf": 24000,
            "construction_type": "medical office",
            "population_projection_2060": 5500000,
            "unemployment_rate": 3.8,
            "national_unemployment": 4.4,
            "absorption_rate_sf": 15000,
            "major_trends": ["hybrid work adoption", "flight to quality", "suburban growth"]
        }}
        """
        
        try:
            response = self.openai_client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": f"You are a commercial real estate market analyst. Provide current, realistic {current_quarter} market data specific to {county}, {state}. Return only valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                response_format={"type": "json_object"},
            )

            content = response.choices[0].message.content.strip()
            # Strip accidental markdown fences before parsing
            if content.startswith("```"):
                content = content.split("```", 2)[1].lstrip("json").strip()
            return json.loads(content)
        except Exception as e:
            logger.error(f"Error generating market data: {e}")
            # Return default data as fallback
            return {
                "quarter": current_quarter,
                "direct_vacancy": 12.71,
                "direct_qoq": "+1.02",
                "direct_yoy": "+1.68",
                "sublease_vacancy": 5.51,
                "sublease_qoq": "-0.39",
                "total_vacancy": 18.22,
                "avg_lease_rate": 24.33,
                "lease_rate_yoy": "-0.04",
                "cap_rate": 6.7,
                "class_a_rate": 27.13,
                "class_b_rate": 22.03,
                "class_c_rate": 19.44,
                "construction_sf": 24000,
                "construction_type": "medical office",
                "population_projection_2060": 5500000,
                "unemployment_rate": 3.8,
                "national_unemployment": 4.4
            }

    def _generate_market_overview(self, context: str, property_type: str, market_data: Dict) -> str:
        """Generate a location-specific market overview via AI (not a static template)."""
        county = self._ctx_value(context, 'County')
        state = self._ctx_value(context, 'State')
        quarter = market_data.get('quarter', self._current_quarter())

        prompt = f"""Write a professional 2-paragraph market overview for the {property_type} market
in {county}, {state}, as of {quarter}.

Ground everything specifically in {county}, {state}. Do NOT reference any other state, county, or city.
Incorporate these current metrics naturally in prose: total vacancy about {market_data.get('total_vacancy', 'n/a')}%,
average asking lease rate about ${market_data.get('avg_lease_rate', 'n/a')}/SF, local unemployment about
{market_data.get('unemployment_rate', 'n/a')}% versus the national average of about {market_data.get('national_unemployment', 'n/a')}%.

Cover local population and employment growth drivers, demand for {property_type} space, and current leasing
dynamics (e.g., sublease availability, flight to quality, hybrid-work effects where relevant).
Plain text only. No markdown, no bullet points, no headings."""
        return self._get_ai_response(prompt)

    def _generate_vacancy_rates(self, context: str, property_type: str, market_data: Dict) -> str:
        """Generate Key Market Metrics bullets in CLIENT-report style (proxy metrics + Note)."""
        county = self._ctx_value(context, "County")
        state = self._ctx_value(context, "State")
        quarter = market_data.get("quarter", self._current_quarter())
        vacancy = market_data.get("total_vacancy", "n/a")
        rent = market_data.get("avg_lease_rate", "n/a")
        construction_sf = market_data.get("construction_sf", 0) or 0
        cap_rate = market_data.get("cap_rate", "n/a")

        prompt = f"""Write the Key Market Metrics block under "Vacancy Rates" for a {property_type}
property in {county}, {state}, as of {quarter}.

Match this EXACT structure and tone (plain text only — no markdown, no **bold** markers):

• Vacancy (metro proxy): {vacancy}% ({quarter}); <one short sentence on metro trend and how {county} typically tracks it>.
• Asking Rents (metro proxy): ~${rent}/SF market asking rent ({quarter}), with brief rent-growth context. Note that {property_type} reporting is not commonly split into Class A/B.
• Construction Pipeline (metro): ~{int(construction_sf):,} SF under construction as of {quarter}; one short pipeline/delivery note.
• Investment Indicators (metro): Average {property_type} capitalization rate ~{cap_rate}% ({quarter}).
Note: Reliable {county}-only breakouts for {property_type} vacancy and asking rents are limited; metro-level metrics are the accepted proxy in most public reports and broker opinions of value.

Rules:
- One bullet per line, each starting with "• " (bullet character + space).
- After the four bullets, one "Note:" paragraph (no bullet).
- Separate each item with a blank line.
- Do NOT output Direct/Sublease/Total vacancy line items.
- Do NOT invent other counties/states. Ground everything in {county}, {state} / its metro.
"""
        return self._get_ai_response(prompt)

    def _generate_lease_rates(self, context: str, property_type: str, market_data: Dict) -> str:
        """Generate lease-rates prose (CLIENT style) — not Class A/B/C bullet lists."""
        county = self._ctx_value(context, "County")
        state = self._ctx_value(context, "State")
        quarter = market_data.get("quarter", self._current_quarter())
        rent = market_data.get("avg_lease_rate", "n/a")
        rent_yoy = market_data.get("lease_rate_yoy", "n/a")

        prompt = f"""Write ONE short paragraph for "Lease Rates (Based on Public Listings & Industry Trends)"
for {property_type} in {county}, {state}, as of {quarter}.

Style example:
"{property_type} lease rates in {county} generally align with broader metro trends, averaging around ${rent}/SF as of {quarter}. Annual rent growth has remained steady{'' if rent_yoy == 'n/a' else f' ({rent_yoy} YoY)'}, supported by disciplined supply and healthy tenant demand in well-located corridors."

Rules:
- Plain text only, single paragraph, no bullets, no Class A/B/C splits, no markdown.
- Keep it to 2–4 sentences. Do not reference other markets.
"""
        return self._get_ai_response(prompt)

    def _generate_construction_activity(self, context: str, property_type: str, market_data: Dict) -> str:
        """Generate a location-specific construction activity section via AI."""
        county = self._ctx_value(context, 'County')
        state = self._ctx_value(context, 'State')
        quarter = market_data.get('quarter', self._current_quarter())
        construction_sf = market_data.get('construction_sf', 0) or 0

        prompt = f"""Write a short "Construction Activity" section for the {property_type} market in
{county}, {state}, as of {quarter}.

Use these figures: approximately {int(construction_sf):,} SF under construction, primarily
{market_data.get('construction_type', property_type)} product.
Reference realistic, plausible recent deliveries and pipeline trends specific to {county}, {state}.
Do NOT reference any other market (no Salt Lake, Lehi, Provo, etc. unless that is the actual county/state).

Return 2-3 concise bullet points, each starting with "• ". Plain text only. Separate bullets with blank lines."""
        return self._get_ai_response(prompt)

    def _generate_market_trends(self, context: str, property_type: str, market_data: Dict) -> str:
        """Generate location-specific market trends via AI."""
        county = self._ctx_value(context, 'County')
        state = self._ctx_value(context, 'State')
        quarter = market_data.get('quarter', self._current_quarter())
        trends = ", ".join(market_data.get('major_trends', []) or [])

        prompt = f"""Write a "Trends & Forecast" section for the {property_type} market in {county}, {state}, as of {quarter}.
{f'Relevant current themes to weave in: {trends}.' if trends else ''}
Cover 3-4 distinct trends (e.g., population/employment growth, leasing flexibility/sublease, supply pipeline,
hybrid-work or e-commerce effects) as they specifically apply to {county}, {state} and to {property_type}.
Do NOT reference any other market.
Format: each trend is one bullet starting with "• " and a short label then colon (e.g. "• Demand skew to small shops: ...").
Separate bullets with blank lines. Plain text only — no markdown (no **, *, #, or backticks)."""
        return self._get_ai_response(prompt)

    def _generate_investment_insights(self, context: str, property_type: str, market_data: Dict) -> str:
        """Generate location-specific investment insights via AI."""
        county = self._ctx_value(context, 'County')
        state = self._ctx_value(context, 'State')

        prompt = f"""Write "Investment Insights" for a {property_type} property in {county}, {state}.
Provide 3 concise, actionable insights grounded in the local {county}, {state} market and {property_type} fundamentals.
Do NOT reference any other market.
Return 3 bullets, each starting with "• " and a short label then colon. Separate with blank lines.
Plain text only — no markdown (no **, *, #, or backticks)."""
        return self._get_ai_response(prompt)

    def _generate_market_recommendations(self, context: str, property_type: str, market_data: Dict) -> str:
        """Generate location-specific recommendations via AI (CLIENT short-bullet style)."""
        county = self._ctx_value(context, "County")
        state = self._ctx_value(context, "State")

        prompt = f"""Write "Recommendations" for a {property_type} Broker Opinion of Value in {county}, {state}.

Return EXACTLY 3 short bullets in this style (plain text only):
• Prioritize <one concrete action grounded in {county} {property_type} conditions>.
• Target <one concrete leasing / tenant / location action>.
• For value-add, focus on <one concrete underwriting / lease-up / mark-to-market action>.

Rules:
- Each bullet is ONE sentence (max ~35 words). Start with "• ".
- Separate bullets with a blank line.
- Do NOT use markdown (no **, no *, no #, no backticks).
- Do NOT label bullets as Investors / Tenants / Developers.
- Do NOT reference any other market.
"""
        return self._get_ai_response(prompt)

    @staticmethod
    def _strip_markdown(text: str) -> str:
        """Remove common markdown markers AI sometimes emits into Word body text."""
        if not text:
            return text or ""
        # Bold / italic
        text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
        text = re.sub(r"__(.+?)__", r"\1", text)
        text = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"\1", text)
        text = re.sub(r"`([^`]+)`", r"\1", text)
        # Headings / list markers that are not our bullets
        text = re.sub(r"(?m)^#{1,6}\s*", "", text)
        return text

    def _generate_data_sources(self, context: str = "", property_type: str = "", market_data: Dict = None) -> str:
        """Generate a current, market-relevant data sources list via AI."""
        market_data = market_data or {}
        county = self._ctx_value(context, 'County')
        state = self._ctx_value(context, 'State')
        quarter = market_data.get('quarter', self._current_quarter())
        year = datetime.now().year

        prompt = f"""Write the "Data Sources & Disclaimer" sources block for a {property_type} Broker Opinion
of Value in {county}, {state}, as of {quarter}.

Match this EXACT structure (plain text only — no markdown):

This analysis relies on multiple public sources:

• CoStar – {county}/{state} metro {property_type} Market {quarter} (vacancy rates, rental rates, absorption)
• CBRE / JLL / similar brokerage – {property_type} Market Report {quarter} (market trends and leasing activity)
• U.S. Census Bureau – latest population, income, and housing statistics for {county}, {state}
• Esri Business Analyst / GeoEnrichment – current demographic and consumer spending data
• U.S. Bureau of Labor Statistics – current employment and unemployment data
• Local county appraisal / assessor records – public property records for {county}

Rules:
- Start with the italic-style intro line exactly: "This analysis relies on multiple public sources:"
- Then 5–6 bullets, each starting with "• " (bullet character). Separate bullets with blank lines.
- Do NOT use a numbered list (no "1.", "2.", "Sources Used:").
- Keep each bullet to one line. Sources must be current ({year} / {quarter}), not older than two years.
- Do NOT invent Utah-specific sources unless the property is in Utah.
"""
        text = self._get_ai_response(prompt)
        # Force CLIENT-style bullets even if the model returns "1. 2. 3."
        text = text.replace("Sources Used:", "This analysis relies on multiple public sources:")
        text = re.sub(r"(?m)^\s*\d+\.\s+", "• ", text)
        return text

    @staticmethod
    def _ctx_value(context: str, label: str) -> str:
        """Safely pull a labelled value (e.g. 'County') out of the context block."""
        try:
            return context.split(f'{label}:')[1].split('\n')[0].strip()
        except (IndexError, AttributeError):
            return ""

    @staticmethod
    def _current_quarter() -> str:
        now = datetime.now()
        return f"Q{(now.month - 1) // 3 + 1} {now.year}"

    def _generate_market_analysis_sections(self, context: str, property_type: str) -> Dict[str, str]:
        """Generate all market analysis sections"""
        
        # Get current market data
        county = context.split('County:')[1].split('\n')[0].strip() if 'County:' in context else 'Utah County'
        state = context.split('State:')[1].split('\n')[0].strip() if 'State:' in context else 'Utah'
        
        market_data = self._search_current_market_data(county, state, property_type)
        
        sections = {
            'market_overview': self._generate_market_overview(context, property_type, market_data),
            'vacancy_rates': self._generate_vacancy_rates(context, property_type, market_data),
            'lease_rates': self._generate_lease_rates(context, property_type, market_data),
            'construction_activity': self._generate_construction_activity(context, property_type, market_data),
            'market_trends': self._generate_market_trends(context, property_type, market_data),
            'investment_insights': self._generate_investment_insights(context, property_type, market_data),
            'market_recommendations': self._generate_market_recommendations(context, property_type, market_data),
            'market_data_sources': self._generate_data_sources(context, property_type, market_data)
        }
        # Word shows literal **bold** if markdown slips through — strip it
        return {k: self._strip_markdown(v) for k, v in sections.items()}

    def generate_comprehensive_content(self, address: str, property_data: PropertyReportData) -> PropertyReportData:
        """
        Generate all content sections using AI with the exact formatting requirements
        """
        logger.info(f"Generating comprehensive content for: {address}")
        
        # Create detailed context for AI
        transaction = ""
        if property_data.lease_or_sale:
            kind = property_data.lease_or_sale.strip().lower()
            if "lease" in kind:
                transaction = "Transaction Context: The property is being evaluated for LEASE."
            else:
                transaction = "Transaction Context: The property is being evaluated for SALE."
        context = f"""
        Property Address: {address}
        Property Type: {property_data.property_type}
        County: {property_data.county}
        State: {property_data.state}
        Coordinates: {property_data.latitude}, {property_data.longitude}
        {transaction}
        """
        
        # Generate property analysis sections
        sections = {
            'property_summary': self._generate_property_summary(context),
            'location_summary': self._generate_location_summary(context),
            'demographic_analysis': self._generate_demographic_analysis(context),
            'size_and_topography': self._generate_size_topography(context),
            'population_analysis': self._generate_population_analysis(context),
            'household_trends': self._generate_household_trends(context),
            'employment_analysis': self._generate_employment_analysis(context),
            'economic_factors': self._generate_economic_factors(context),
            'community_services': self._generate_community_services(context),
            'swot_analysis': self._generate_swot_analysis(context)
        }
        
        # Update property data with generated content
        property_data.property_summary = sections['property_summary']
        property_data.location_summary = sections['location_summary']
        property_data.demographic_analysis = sections['demographic_analysis']
        property_data.size_and_topography = sections['size_and_topography']
        property_data.population_analysis = sections['population_analysis']
        property_data.household_trends = sections['household_trends']
        property_data.employment_analysis = sections['employment_analysis']
        property_data.economic_factors = sections['economic_factors']
        property_data.community_services = sections['community_services']
        
        # SWOT Analysis
        swot = sections['swot_analysis']
        property_data.swot_strengths = swot.get('strengths', '')
        property_data.swot_weaknesses = swot.get('weaknesses', '')
        property_data.swot_opportunities = swot.get('opportunities', '')
        property_data.swot_threats = swot.get('threats', '')
        
        # Generate market analysis sections
        market_sections = self._generate_market_analysis_sections(context, property_data.property_type)
        
        # Update property data with market analysis
        property_data.market_overview = market_sections['market_overview']
        property_data.vacancy_rates = market_sections['vacancy_rates']
        property_data.lease_rates = market_sections['lease_rates']
        property_data.construction_activity = market_sections['construction_activity']
        property_data.market_trends = market_sections['market_trends']
        property_data.investment_insights = market_sections['investment_insights']
        property_data.market_recommendations = market_sections['market_recommendations']
        property_data.market_data_sources = market_sections['market_data_sources']

        comp_context = self._format_comp_context(property_data.comps)

        # Generate additional BOV narrative sections
        property_data.executive_summary = self._get_ai_response(
            f"Write a concise 1-paragraph executive summary for a Broker Opinion of Value of a "
            f"{property_data.property_type} property. Context:\n{context}\n"
            f"{comp_context}"
            f"Summarize location, key value drivers, and the value conclusion. Plain text only.",
        )
        property_data.regional_analysis = self._shorten_regional_analysis(
            self._get_ai_response(
                f"Write a SHORT but worthy REGIONAL ANALYSIS for the subject in "
                f"{property_data.county}, {property_data.state} ({datetime.now().year}).\n\n"
                f"Match this style (plain text, no markdown, no bullets):\n"
                f"Paragraph 1: population growth/size and household/income character that support "
                f"{property_data.property_type} demand (2 sentences max).\n"
                f"Paragraph 2: key employers/industries plus major roads/airport access and why that "
                f"helps the local {property_data.property_type} market (2 sentences max).\n\n"
                f"HARD LIMITS: exactly 2 paragraphs, blank line between them, 85-100 words TOTAL. "
                f"Name concrete employers/roads when plausible. Never mention another county/state.\n"
                f"Context:\n{context}",
            )
        )
        # CLIENT layout: SALES CONCLUSION heading + OPINIONS OF VALUE table only
        # (no long narrative between/after those elements)
        property_data.sales_conclusion = ""
        property_data.reconciliation_summary = self._get_ai_response(
            f"Write 2 short sentences for the RECONCILIATION TABLE narrative (above the valuation grid). "
            f"Mention the comparable $/SF range when comps are available and that the sales comparison "
            f"approach supports the opinion of value for this {property_data.property_type}. "
            f"Context:\n{context}\n{comp_context}"
            f"Plain text only. Keep under 60 words.",
        )
        # Short NOTES cell for the valuation grid (CLIENT style one-liner)
        property_data.reconciliation_notes = self._build_reconciliation_notes(property_data)

        return property_data

    @staticmethod
    def _format_comp_context(comps: List[Any]) -> str:
        if not comps:
            return ""
        lines = ["Comparable sales from uploaded CoStar PDF:"]
        for comp in sorted(comps, key=lambda c: getattr(c, "comp_number", 0))[:6]:
            lines.append(
                f"- Comp {getattr(comp, 'comp_number', '?')}: {getattr(comp, 'address', '')}, "
                f"Sale {getattr(comp, 'sale_price', 'N/A')}, "
                f"{getattr(comp, 'sale_price_sf', 'N/A')}/SF, "
                f"{getattr(comp, 'comp_sf', '')} SF"
            )
        return "\n".join(lines) + "\n"

    def _build_reconciliation_notes(self, property_data: "PropertyReportData") -> str:
        """One-line NOTES cell for the valuation grid (CLIENT style)."""
        comps = getattr(property_data, "comps", None) or []
        n_comps = len(comps) if comps else 0

        rounded = None
        # Prefer finalized valuation from table_values when available
        tv = getattr(property_data, "table_values", None) or {}
        rounded = tv.get("{{market_value_rounded}}")
        if not rounded and getattr(property_data, "bov_dataset", None):
            val = (property_data.bov_dataset or {}).get("valuation") or {}
            mv = val.get("market_value_rounded") or val.get("market_value")
            if mv is not None:
                try:
                    rounded = f"${int(round(float(mv))):,}"
                except (TypeError, ValueError):
                    rounded = str(mv)

        if not rounded:
            rounded = "the concluded market value"

        if n_comps > 0:
            return (
                f"The sales comparison approach yields a value of {rounded} "
                f"based on the average $/SF from {n_comps} comparables."
            )
        return (
            f"The sales comparison approach yields a value of {rounded} "
            f"based on comparable sales analysis."
        )

    @staticmethod
    def _shorten_regional_analysis(text: str, max_words: int = 100) -> str:
        """Keep regional analysis to ~2 short paragraphs (CLIENT density)."""
        if not text:
            return ""
        text = ComprehensivePropertyReportGenerator._strip_markdown(text).strip()
        parts = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
        if len(parts) == 1:
            sentences = re.split(r"(?<=[.!?])\s+", parts[0])
            sentences = [s.strip() for s in sentences if s.strip()]
            if len(sentences) >= 2:
                mid = max(1, len(sentences) // 2)
                parts = [" ".join(sentences[:mid]), " ".join(sentences[mid:])]
            else:
                parts = [parts[0]]
        parts = parts[:2]

        def _clip(paragraph: str, budget: int) -> str:
            words = paragraph.split()
            if len(words) <= budget:
                return paragraph
            clipped = " ".join(words[:budget]).rstrip(",;:")
            if not clipped.endswith((".", "!", "?")):
                clipped += "."
            return clipped

        budget_each = max(35, max_words // max(len(parts), 1))
        parts = [_clip(p, budget_each) for p in parts]
        joined = "\n\n".join(parts)
        words = joined.split()
        if len(words) > max_words:
            joined = " ".join(words[:max_words]).rstrip(",;:")
            if not joined.endswith((".", "!", "?")):
                joined += "."
        return joined

    def _style_regional_analysis(self, doc: Document) -> None:
        """CLIENT style: italic blue body under REGIONAL ANALYSIS."""
        from docx.shared import RGBColor, Pt

        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip().upper() != "REGIONAL ANALYSIS":
                continue
            for j in range(i + 1, len(doc.paragraphs)):
                nxt = doc.paragraphs[j]
                style_name = (nxt.style.name if nxt.style else "") or ""
                upper = nxt.text.strip().upper()
                if "Heading" in style_name or upper in (
                    "DEMOGRAPHIC ANALYSIS",
                    "LOCATION SUMMARY",
                    "PROPERTY SUMMARY",
                ):
                    break
                if not nxt.text.strip():
                    continue
                for run in nxt.runs:
                    run.italic = True
                    run.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)
                    run.font.size = Pt(12)
            break

    def _generate_property_summary(self, context: str) -> str:
        """Generate property summary matching the exact format"""
        example = """The subject is located in Salem, in Utah County. It is part of the Provo-Orem MSA. The subject property is located in northern Utah within the official boundaries of Utah County. The county is situated directly south of Salt Lake County. This area is generally called the Provo/Orem metropolitan area and is approximately 45 miles south of metropolitan Salt Lake, which is the financial center for the Intermountain Region. This region encompasses all of Utah, southern Idaho, southwestern Wyoming, and eastern Nevada. Utah County is part of a four-county area that is commonly known as the Wasatch Front. Provo is the Utah County seat."""
        
        prompt = f"""
        Generate a property summary paragraph following this EXACT format and style:
        {example}
        
        Use this context: {context}
        
        Maintain the same structure: location, county, MSA/metropolitan area, regional context, broader region description, and county seat information.
        """
        
        return self._get_ai_response(prompt)

    def _generate_location_summary(self, context: str) -> str:
        """Generate location summary matching the exact format"""
        example = """The subject is located on the corner of Hwy 198 and Elk Ridge Drive with good access and exposure. A major thoroughfare in the area is Hwy 198 which partially fronts the subject. The location also offers very close proximity to Salem Pond, Salem High School, Salem Community Center, Salem City Recreation, with limited retail areas in close proximity. The subject is surrounded by vacant land and residential uses. Utah County is broken up into three sectors.  North County (Lindon to Lehi) Central County (Provo/Orem) and South County (Springville to Payson). Central county accounts for a lot of the class B office buildings. The following is taken from Reis, it shows the market area with an arrow pointing to the subject."""
        
        prompt = f"""
        Generate a DETAILED location summary following this FORMAT and style (adapt fully to the real location):
        {example}

        Use this context: {context}

        Write a thorough 6-9 sentence description. Include, when applicable to the actual location:
        - the specific corridor/intersection and major thoroughfares fronting or serving the site
        - access and visibility/exposure characteristics
        - nearby anchors, retail, employers, schools, parks, and other notable amenities (name plausible real ones for that area)
        - surrounding land uses and development character
        - how the site sits within the broader county/metro and its submarket subdivisions
        Match the older template's level of detail. Plain text only, no markdown.
        """

        return self._get_ai_response(prompt)

    def _generate_demographic_analysis(self, context: str) -> str:
        """Generate demographic analysis - note this matches property_summary in the example"""
        return self._generate_property_summary(context)

    def _generate_size_topography(self, context: str) -> str:
        """Generate size and topography description"""
        example = """The surrounding mountains form a valley about 30 miles wide and 50 miles long. Utah Lake is located centrally to the valley and is Utah's largest freshwater lake. The Wasatch Mountains, which provide a beautiful background to the county on the east, nearly converge with Utah Lake on the west to form the southern boundary south of Santaquin City. The northern boundary is considered the "point of the mountain" which is just north of Lehi City. The elevation varies from 4,480 to 11,928 feet (Mt. Nebo) above sea level. Utah Lake and Mt. Timpanogos present a mountainous scenic backdrop within this metropolitan setting."""
        
        prompt = f"""
        Generate a size and topography paragraph following this EXACT format and style:
        {example}
        
        Use this context: {context}
        
        Include: geographical features, valley dimensions, major landmarks, elevation ranges, and scenic elements.
        """
        
        return self._get_ai_response(prompt)

    def _generate_population_analysis(self, context: str) -> str:
        """Generate population analysis with specific statistics"""
        example = """According to Pitney Bowes/Gadberry Group - GroundView®, a Geographic Information System (GIS) Company, Utah County had a 2020 total population of 649,258 and experienced an annual growth rate of 2.3%, which was higher than the Utah annual growth rate of 1.6%. The county accounted for 20.0% of the total Utah population (3,254,284). Within the county the population density was 304 people per square mile compared to the lower Utah population density of 38 people per square mile and the lower United States population density of 92 people per square mile."""
        
        prompt = f"""
        Generate a population analysis paragraph following this EXACT format and style:
        {example}
        
        Use this context: {context}
        
        Include: data source attribution, specific population numbers, growth rates, state comparisons, population density comparisons. Use current 2024 data where available.
        """
        
        return self._get_ai_response(prompt)

    def _generate_household_trends(self, context: str) -> str:
        """Generate household trends analysis"""
        example = """The 2020 number of households in the county was 178,689. The number of households in the county is projected to grow by 2.0% annually, increasing the number of households to 197,669 by 2025. The 2020 average household size for the county was 3.55, which was 37.61% larger than the United States average household size of 2.58 for 2020. The average household size in the county is anticipated to retract by 0.06% annually, reducing the average household size to 3.54 by 2025."""
        
        prompt = f"""
        Generate a household trends paragraph following this EXACT format and style:
        {example}
        
        Use this context: {context}
        
        Include: specific household numbers, growth projections, average household size, national comparisons, and future projections. Use current 2024 data and project to 2029.
        """
        
        return self._get_ai_response(prompt)

    def _generate_employment_analysis(self, context: str) -> str:
        """Generate employment analysis using the most recent years available."""
        year = datetime.now().year
        example = f"""Total employment has increased annually over the past three years in the state by roughly 2.5% and by roughly 3.9% in the county. From {year-1} to {year} unemployment fell in the state by about 0.4% and by about 0.4% in the county, with the county's rate remaining below the national average. Over the most recent month, unemployment declined by about 0.3% at both the state and county level."""

        prompt = f"""
        Generate an employment analysis paragraph following this FORMAT and style (do not copy its numbers):
        {example}

        Use this context: {context}

        Requirements:
        - Use the MOST RECENT data available (reference {year-2}-{year}); never cite a decade-old range such as 2010-2019.
        - Include: recent employment growth rates, unemployment trends, state vs county comparisons, and the latest month-over-month change.
        - Base facts on the specific state and county in the context. Plain text only.
        """

        return self._get_ai_response(prompt)

    def _generate_economic_factors(self, context: str) -> str:
        """Generate economic factors analysis"""
        example = """Salem is a suburb of Payson and Provo/Orem market area. Salem is still considered somewhat of a rural area but over the years has begun to be built out. A majority of resident's commute to other cities within the metropolitan area for employment. The largest industries in the city include manufacturing, public administration agricultural uses and retail trade. The local economy consists of commercial and industrial businesses on the main arterials. The city's commercial area is on Hwy 198, featuring retail, office, residential, and financial services."""
        
        prompt = f"""
        Generate an economic factors paragraph following this EXACT format and style:
        {example}
        
        Use this context: {context}
        
        Include: suburban/rural character, development patterns, commuting patterns, major industries, commercial areas, and business types.
        """
        
        return self._get_ai_response(prompt)

    def _generate_community_services(self, context: str) -> str:
        """Generate community services description"""
        example = """Community services and facilities are readily available in the surrounding area. These include public services such as fire stations, hospitals, police stations, and schools (all ages). GreatSchools.org is an on-line tool that rates every school on a scale of one to ten based on test scores. They also track parents rating of the school on a one to five scale."""
        
        prompt = f"""
        Generate a community services paragraph following this EXACT format and style:
        {example}
        
        Use this context: {context}
        
        Include: availability of services, specific service types, educational resources, and rating systems.
        """
        
        return self._get_ai_response(prompt)

    def _generate_swot_analysis(self, context: str) -> Dict[str, str]:
        """Generate SWOT analysis components"""
        prompt = f"""
        Generate a SWOT analysis for this property location. Return as JSON with keys: strengths, weaknesses, opportunities, threats.
        
        Context: {context}
        
        Examples:
        - Strengths: "Easy access from Highway 198 Within close proximity to residential developments"
        - Weaknesses: "The subject has average to weak visibility"  
        - Opportunities: "Opportunity for development of improvement on property"
        - Threats: "There is excess land around the property that could possibility be developed."
        
        Keep each section short and concise, focusing on location-specific factors.
        """
        
        response = self._get_ai_response(prompt, json_response=True)
        try:
            return json.loads(response)
        except:
            return {
                'strengths': 'Good location with development potential',
                'weaknesses': 'Limited visibility from main roads',
                'opportunities': 'Opportunity for future development',
                'threats': 'Competition from nearby available land'
            }

    def _get_ai_response(self, prompt: str, json_response: bool = False) -> str:
        """Get response from OpenAI API"""
        try:
            kwargs = {
                "model": "gpt-4o",
                "messages": [
                    {"role": "system", "content": "You are a professional commercial real estate analyst. Use the example only for FORMAT/style; base all facts on the specific location provided in the context. Never reference Utah, Salt Lake, Provo, or Lehi unless the subject property is actually there. Provide factual, current information." + (" Respond in valid JSON format." if json_response else "")},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.3,
            }
            if json_response:
                kwargs["response_format"] = {"type": "json_object"}
            response = self.openai_client.chat.completions.create(**kwargs)
            return response.choices[0].message.content.strip()
        except Exception as e:
            logger.error(f"Error getting AI response: {e}")
            return "Content generation failed"

    def create_property_report(self, 
                             address: str,
                             prepared_by: str = "",
                             prepared_by_title: str = "",
                             prepared_by_company: str = "",
                             prepared_by_address: str = "",
                             prepared_for: str = "",
                             prepared_for_title: str = "",
                             prepared_for_company: str = "",
                             prepared_for_address: str = "",
                             property_name: str = "",
                             property_type: str = "Vacant Land",
                             **kwargs) -> PropertyReportData:
        """
        Create complete property report data
        """
        logger.info(f"Creating property report for: {address}")

        comps = kwargs.pop("comps", None) or []
        
        # Get coordinates and location details
        lat, lng, address_details = self.get_coordinates_and_details(address)
        
        # Initialize property data
        property_data = PropertyReportData(
            address=address,
            prepared_by=prepared_by,
            prepared_by_title=prepared_by_title,
            prepared_by_company=prepared_by_company,
            prepared_by_address=prepared_by_address,
            prepared_for=prepared_for,
            prepared_for_title=prepared_for_title,
            prepared_for_company=prepared_for_company,
            prepared_for_address=prepared_for_address,
            property_name=property_name or f"{address_details.get('city', 'Property')} {property_type}",
            property_type=(property_type or "").strip().title(),
            state=address_details.get('state', ''),
            county=address_details.get('county', ''),
            latitude=str(lat),
            longitude=str(lng),
            **kwargs
        )
        
        # Get property images
        aerial_path, street_view_path = self.get_property_images(address, lat, lng)
        property_data.aerial_image_path = aerial_path
        property_data.street_view_image_path = street_view_path
        property_data.comps = comps
        
        # Generate all content sections (including market analysis)
        property_data = self.generate_comprehensive_content(address, property_data)

        # Build BOV table values (Esri GeoEnrichment first, AI fallback)
        try:
            lat_f = float(property_data.latitude)
            lng_f = float(property_data.longitude)
        except (TypeError, ValueError):
            lat_f, lng_f = 0.0, 0.0
        property_data.table_values = self.build_bov_dataset(
            lat_f, lng_f, property_data
        )
        # Rebuild short valuation NOTES after GBA/valuation math is final
        property_data.reconciliation_notes = self._build_reconciliation_notes(property_data)
        property_data.table_values["{{reconciliation_notes}}"] = property_data.reconciliation_notes
        property_data.table_values["{{reconciliation_summary}}"] = (
            property_data.reconciliation_summary or ""
        )
        property_data.table_values["{{sales_conclusion}}"] = (
            property_data.sales_conclusion or ""
        )

        return property_data

    def build_bov_dataset(self, lat: float, lng: float, property_data: "PropertyReportData") -> Dict[str, str]:
        """Build the full set of BOV table placeholder values.

        Strategy: generate a complete realistic dataset with AI, then overlay
        real Esri GeoEnrichment numbers where the Esri key is available.
        """
        county = property_data.county or "the county"
        state = property_data.state or "the state"
        property_type = property_data.property_type or "Commercial"

        dataset = self._generate_bov_demographics_ai(county, state, property_type)

        # Use the entered GBA (Gross Building Area) to drive the valuation math
        gba = self._to_number(property_data.lot_area)
        if gba:
            valuation = dataset.setdefault("valuation", {})
            price_psf = self._to_number(valuation.get("price_psf")) or 265.0
            valuation["building_sf"] = gba
            valuation["price_psf"] = price_psf
            market_value = price_psf * gba
            valuation["market_value"] = market_value
            valuation["market_value_rounded"] = round(market_value / 10000) * 10000
            valuation["value_aggressive"] = round(market_value * 1.04)
            valuation["value_conservative"] = round(market_value * 0.96)

        # Overlay real Esri ring demographics when available
        try:
            rings = self.location_service.get_ring_demographics(lat, lng)
            if rings:
                self._overlay_esri_rings(dataset, rings)
                dataset["demographics_source"] = "Esri GeoEnrichment"
                logger.info("Overlaid real Esri ring demographics onto BOV dataset")
        except Exception as exc:
            logger.warning("Could not overlay Esri ring demographics: %s", exc)

        # Overlay US / State / County tables from Esri admin geographies (not point buffer)
        try:
            admin = self.location_service.get_admin_demographics(lat, lng)
            if admin:
                self._overlay_esri_admin(dataset, admin)
                dataset["demographics_source"] = "Esri GeoEnrichment"
                logger.info(
                    "Overlaid Esri admin demographics (us=%s state=%s county=%s)",
                    bool(admin.get("us")),
                    bool(admin.get("state")),
                    bool(admin.get("county")),
                )
        except Exception as exc:
            logger.warning("Could not overlay Esri admin demographics: %s", exc)

        property_data.bov_dataset = dataset
        return self._format_bov_placeholders(dataset, property_data)

    def _overlay_esri_admin(self, dataset: Dict, admin: Dict[str, Dict]) -> None:
        """Replace US/State/County population, density, HH, tenure from Esri admin levels."""
        pop = dataset.setdefault("population", {})
        density = dataset.setdefault("density", {})
        households = dataset.setdefault("households", {})
        hh_size = dataset.setdefault("hh_size", {})
        tenure = dataset.setdefault("tenure", {})

        for year in ("2010", "2020", "2025"):
            pop.setdefault(year, {})
        for year in ("2020", "2025"):
            density.setdefault(year, {})
        for key in ("2024", "2029", "cagr"):
            households.setdefault(key, {})
            hh_size.setdefault(key, {})
        tenure.setdefault("owner", {})
        tenure.setdefault("renter", {})

        for geo in ("us", "state", "county"):
            row = admin.get(geo) or {}
            if not row:
                continue

            if row.get("pop_2010") is not None:
                pop["2010"][geo] = row["pop_2010"]
            if row.get("pop_2020") is not None:
                pop["2020"][geo] = row["pop_2020"]
            if row.get("pop_cy") is not None:
                pop["2025"][geo] = row["pop_cy"]

            if row.get("density_2020") is not None:
                density["2020"][geo] = row["density_2020"]
            if row.get("density_cy") is not None:
                density["2025"][geo] = row["density_cy"]

            if row.get("hh_cy") is not None:
                households["2024"][geo] = row["hh_cy"]
            if row.get("hh_fy") is not None:
                households["2029"][geo] = row["hh_fy"]
            cagr_hh = self._cagr(row.get("hh_cy"), row.get("hh_fy"), years=5)
            if cagr_hh is not None:
                households["cagr"][geo] = cagr_hh

            if row.get("hhsize_cy") is not None:
                hh_size["2024"][geo] = row["hhsize_cy"]
            hhsize_fy = row.get("hhsize_fy")
            if hhsize_fy is None and row.get("pop_fy") and row.get("hh_fy"):
                hhsize_fy = round(float(row["pop_fy"]) / float(row["hh_fy"]), 2)
            if hhsize_fy is not None:
                hh_size["2029"][geo] = hhsize_fy
            cagr_hs = self._cagr(row.get("hhsize_cy"), hhsize_fy, years=5)
            if cagr_hs is not None:
                hh_size["cagr"][geo] = cagr_hs

            if row.get("owner_pct") is not None:
                tenure["owner"][geo] = row["owner_pct"]
            if row.get("renter_pct") is not None:
                tenure["renter"][geo] = row["renter_pct"]

    @staticmethod
    def _cagr(start, end, years: int = 5):
        """Compound annual growth rate (%) over `years`."""
        try:
            start_f = float(start)
            end_f = float(end)
            if start_f <= 0 or end_f <= 0 or years <= 0:
                return None
            return round(((end_f / start_f) ** (1.0 / years) - 1.0) * 100.0, 1)
        except (TypeError, ValueError, ZeroDivisionError):
            return None

    def _overlay_esri_county(self, dataset: Dict, county_demo: Dict) -> None:
        """Legacy point-level overlay — employment only (never county population).

        Point enrichment returns block-group scale population (~thousands), which
        must not be written into the County column of US/State/County tables.
        """
        emp = dataset.setdefault("employment", {})
        emp.setdefault("total_employment", {})
        emp.setdefault("unemployment_rate", {})

        if county_demo.get("employment_count") is not None:
            emp["total_employment"]["county"] = int(county_demo["employment_count"])
        if county_demo.get("unemployment_rate") is not None:
            emp["unemployment_rate"]["county"] = float(county_demo["unemployment_rate"])

    @staticmethod
    def _to_number(value):
        """Parse a number from strings like '24,500', '708711', or '24500 SF'."""
        if value is None:
            return None
        try:
            text = str(value).strip()
            if not text:
                return None
            # Strip currency / area unit suffixes users often type into GBA
            text = text.replace("$", "").replace(",", "")
            for suffix in (
                "sq. ft.", "sq ft", "sqft", "s.f.", "sf", "acres", "acre", "gba",
            ):
                if text.lower().endswith(suffix):
                    text = text[: -len(suffix)].strip()
                    break
            return float(text)
        except (TypeError, ValueError):
            return None

    def _generate_bov_demographics_ai(self, county: str, state: str, property_type: str) -> Dict:
        """Generate a complete demographic + valuation dataset as JSON via AI."""
        prompt = f"""
        Generate realistic current demographic, employment, and valuation data for a {property_type}
        property in {county}, {state}. Base values on plausible US Census, Esri, and BLS figures.
        Use exactly 10 consecutive years ending with the most recent available year
        (e.g. {datetime.now().year - 9}–{datetime.now().year}) for employment_history —
        never leave gaps and never use a static 2010-2019 sample.

        Return ONLY valid JSON with this exact structure (numbers as plain integers/decimals,
        no commas, no $ signs):
        {{
          "population": {{
            "2010": {{"us": 308745538, "state": 25145561, "county": 2368139}},
            "2020": {{"us": 331449281, "state": 29145505, "county": 2613539}},
            "2025": {{"us": 340110988, "state": 30500000, "county": 2640000}}
          }},
          "density": {{
            "2020": {{"us": 92, "state": 108, "county": 2985}},
            "2025": {{"us": 94, "state": 113, "county": 2982}}
          }},
          "households": {{
            "2024": {{"us": 131000000, "state": 11000000, "county": 982000}},
            "2029": {{"us": 138000000, "state": 12500000, "county": 1037000}},
            "cagr": {{"us": 1.4, "state": 2.6, "county": 1.1}}
          }},
          "hh_size": {{
            "2024": {{"us": 2.52, "state": 2.66, "county": 2.62}},
            "2029": {{"us": 2.42, "state": 2.50, "county": 2.46}},
            "cagr": {{"us": -0.8, "state": -1.3, "county": -1.2}}
          }},
          "tenure": {{
            "owner": {{"us": 65.0, "state": 62.6, "county": 50.8}},
            "renter": {{"us": 35.0, "state": 37.4, "county": 49.2}}
          }},
          "rings": {{
            "1": {{
              "pop_2010": 11800, "pop_2020": 13500, "pop_2024": 13200, "pop_2029": 13100,
              "hh_2010": 4300, "hh_2020": 4500, "hh_2024": 4600, "hh_2029": 4800,
              "avg_hh_income": 150000, "avg_hh_income_2029": 152000,
              "median_hh_income": 120000, "median_hh_income_2029": 122000,
              "per_capita_income": 65000, "per_capita_income_2029": 66000,
              "owner_pct": 60.5, "renter_pct": 39.5,
              "avg_home_value": 550000, "median_home_value": 520000,
              "inc_lt_15k": 180, "inc_15_25": 150, "inc_25_35": 200, "inc_35_50": 350,
              "inc_50_75": 550, "inc_75_100": 480, "inc_100_150": 900, "inc_150_200": 520, "inc_200_plus": 1270,
              "built_2020_later": 120, "built_2010_2019": 480, "built_2000_2009": 620, "built_1990_1999": 540,
              "built_1980_1989": 700, "built_1970_1979": 650, "built_1960_1969": 480, "built_1950_1959": 390,
              "built_1940_1949": 140, "built_1939_earlier": 150,
              "units_1_det": 2800, "units_1_att": 220, "units_2": 90, "units_3_4": 180, "units_5_9": 210,
              "units_10_19": 260, "units_20_49": 240, "units_50_plus": 400, "units_mobile": 40, "units_other": 10
            }},
            "3": {{
              "pop_2010": 78000, "pop_2020": 93000, "pop_2024": 95000, "pop_2029": 96000,
              "hh_2010": 31500, "hh_2020": 36000, "hh_2024": 39000, "hh_2029": 42000,
              "avg_hh_income": 140000, "avg_hh_income_2029": 142000,
              "median_hh_income": 99000, "median_hh_income_2029": 101000,
              "per_capita_income": 58000, "per_capita_income_2029": 59000,
              "owner_pct": 41.7, "renter_pct": 58.3,
              "avg_home_value": 480000, "median_home_value": 450000,
              "inc_lt_15k": 1800, "inc_15_25": 1600, "inc_25_35": 2200, "inc_35_50": 3600,
              "inc_50_75": 5200, "inc_75_100": 4500, "inc_100_150": 7800, "inc_150_200": 4200, "inc_200_plus": 8100,
              "built_2020_later": 1400, "built_2010_2019": 4200, "built_2000_2009": 5100, "built_1990_1999": 4800,
              "built_1980_1989": 6200, "built_1970_1979": 5800, "built_1960_1969": 4100, "built_1950_1959": 3400,
              "built_1940_1949": 1200, "built_1939_earlier": 1300,
              "units_1_det": 22000, "units_1_att": 2100, "units_2": 900, "units_3_4": 1800, "units_5_9": 2400,
              "units_10_19": 3100, "units_20_49": 2800, "units_50_plus": 4500, "units_mobile": 350, "units_other": 80
            }},
            "5": {{
              "pop_2010": 179000, "pop_2020": 210000, "pop_2024": 218000, "pop_2029": 222000,
              "hh_2010": 69000, "hh_2020": 80000, "hh_2024": 86000, "hh_2029": 93000,
              "avg_hh_income": 140000, "avg_hh_income_2029": 143000,
              "median_hh_income": 103000, "median_hh_income_2029": 105000,
              "per_capita_income": 55000, "per_capita_income_2029": 56500,
              "owner_pct": 46.1, "renter_pct": 53.9,
              "avg_home_value": 460000, "median_home_value": 430000,
              "inc_lt_15k": 4200, "inc_15_25": 3800, "inc_25_35": 5100, "inc_35_50": 8200,
              "inc_50_75": 11800, "inc_75_100": 10200, "inc_100_150": 16800, "inc_150_200": 9200, "inc_200_plus": 16900,
              "built_2020_later": 3200, "built_2010_2019": 9800, "built_2000_2009": 11800, "built_1990_1999": 11000,
              "built_1980_1989": 14200, "built_1970_1979": 13200, "built_1960_1969": 9500, "built_1950_1959": 7800,
              "built_1940_1949": 2800, "built_1939_earlier": 3000,
              "units_1_det": 52000, "units_1_att": 4800, "units_2": 2100, "units_3_4": 4200, "units_5_9": 5600,
              "units_10_19": 7200, "units_20_49": 6500, "units_50_plus": 10500, "units_mobile": 900, "units_other": 180
            }}
          }},
          "employment": {{
            "total_employment": {{"us": 161000000, "state": 14500000, "county": 1350000}},
            "unemployment_rate": {{"us": 4.1, "state": 4.0, "county": 3.8}}
          }},
          "employment_history": [
            {{"year": 2016, "state_emp": 11800000, "state_emp_yoy": null, "state_unemp": 4.6,
              "county_emp": 1120000, "county_emp_yoy": null, "county_unemp": 4.0,
              "us_emp": 144000000, "us_unemp": 4.9}},
            {{"year": 2017, "state_emp": 12050000, "state_emp_yoy": 2.1, "state_unemp": 4.3,
              "county_emp": 1145000, "county_emp_yoy": 2.2, "county_unemp": 3.8,
              "us_emp": 146500000, "us_unemp": 4.4}},
            {{"year": 2018, "state_emp": 12300000, "state_emp_yoy": 2.1, "state_unemp": 3.9,
              "county_emp": 1170000, "county_emp_yoy": 2.2, "county_unemp": 3.6,
              "us_emp": 149000000, "us_unemp": 3.9}},
            {{"year": 2019, "state_emp": 12550000, "state_emp_yoy": 2.0, "state_unemp": 3.5,
              "county_emp": 1195000, "county_emp_yoy": 2.1, "county_unemp": 3.4,
              "us_emp": 151000000, "us_unemp": 3.7}},
            {{"year": 2020, "state_emp": 12200000, "state_emp_yoy": -2.8, "state_unemp": 7.6,
              "county_emp": 1160000, "county_emp_yoy": -2.9, "county_unemp": 7.3,
              "us_emp": 142000000, "us_unemp": 8.1}},
            {{"year": 2021, "state_emp": 12600000, "state_emp_yoy": 3.3, "state_unemp": 5.7,
              "county_emp": 1200000, "county_emp_yoy": 3.4, "county_unemp": 5.4,
              "us_emp": 148000000, "us_unemp": 5.4}},
            {{"year": 2022, "state_emp": 13000000, "state_emp_yoy": 3.2, "state_unemp": 4.0,
              "county_emp": 1240000, "county_emp_yoy": 3.3, "county_unemp": 3.8,
              "us_emp": 153000000, "us_unemp": 3.6}},
            {{"year": 2023, "state_emp": 13300000, "state_emp_yoy": 2.3, "state_unemp": 3.9,
              "county_emp": 1270000, "county_emp_yoy": 2.4, "county_unemp": 3.7,
              "us_emp": 155500000, "us_unemp": 3.6}},
            {{"year": 2024, "state_emp": 13600000, "state_emp_yoy": 2.3, "state_unemp": 4.0,
              "county_emp": 1300000, "county_emp_yoy": 2.4, "county_unemp": 3.8,
              "us_emp": 158000000, "us_unemp": 3.9}},
            {{"year": 2025, "state_emp": 13850000, "state_emp_yoy": 1.8, "state_unemp": 3.9,
              "county_emp": 1325000, "county_emp_yoy": 1.9, "county_unemp": 3.6,
              "us_emp": 160000000, "us_unemp": 4.0}}
          ],
          "valuation": {{
            "price_psf": 265.00, "building_sf": 24500,
            "market_value": 6492500, "market_value_rounded": 6490000,
            "value_aggressive": 6752200, "value_conservative": 6232800
          }}
        }}
        """
        try:
            response = self.openai_client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a real estate data analyst. Return only valid JSON, no markdown."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.1,
                response_format={"type": "json_object"},
            )
            return json.loads(response.choices[0].message.content)
        except Exception as exc:
            logger.error("BOV demographics AI generation failed: %s", exc)
            return {}

    def _overlay_esri_rings(self, dataset: Dict, rings: Dict[str, Dict]) -> None:
        """Replace AI ring numbers with real Esri values where present."""
        dataset.setdefault("rings", {})
        esri_map = {
            "pop_2010": ("TOTPOP10", "POP10"),
            "pop_2020": ("TOTPOP20", "POP20", "TOTPOP_CY_PREV"),
            "pop_2024": ("TOTPOP_CY", "TOTPOP_FY"),
            "pop_2029": ("TOTPOP_FY",),
            "hh_2010": ("TOTHH10", "HH10"),
            "hh_2020": ("TOTHH20", "HH20"),
            "hh_2024": ("TOTHH_CY",),
            "hh_2029": ("TOTHH_FY",),
            "avg_hh_income": ("AVGHINC_CY", "AVGHHINC_CY"),
            "avg_hh_income_2029": ("AVGHINC_FY", "AVGHHINC_FY"),
            "median_hh_income": ("MEDHINC_CY",),
            "median_hh_income_2029": ("MEDHINC_FY",),
            "per_capita_income": ("PCI_CY",),
            "per_capita_income_2029": ("PCI_FY",),
            "owner_pct": ("OWNERPCT_CY", "OWNER_CY"),
            "renter_pct": ("RENTERPCT_CY", "RENTER_CY"),
            "avg_home_value": ("AVGVAL_CY", "AVGHOMEVAL_CY"),
            "median_home_value": ("MEDVAL_CY", "MEDHOMEVAL_CY"),
            # Households by income (counts)
            "inc_lt_15k": ("HINC0_CY", "ACSINC0"),
            "inc_15_25": ("HINC15_CY", "ACSINC15"),
            "inc_25_35": ("HINC25_CY", "ACSINC25"),
            "inc_35_50": ("HINC35_CY", "ACSINC35"),
            "inc_50_75": ("HINC50_CY", "ACSINC50"),
            "inc_75_100": ("HINC75_CY", "ACSINC75"),
            "inc_100_150": ("HINC100_CY", "ACSINC100"),
            "inc_150_200": ("HINC150_CY", "ACSINC150"),
            "inc_200_plus": ("HINC200_CY", "ACSINC200"),
            # Year built
            "built_2020_later": ("ACSYB2020", "YB2020_CY"),
            "built_2010_2019": ("ACSYB2010", "YB2010_CY"),
            "built_2000_2009": ("ACSYB2000", "YB2000_CY"),
            "built_1990_1999": ("ACSYB1990", "YB1990_CY"),
            "built_1980_1989": ("ACSYB1980", "YB1980_CY"),
            "built_1970_1979": ("ACSYB1970", "YB1970_CY"),
            "built_1960_1969": ("ACSYB1960", "YB1960_CY"),
            "built_1950_1959": ("ACSYB1950", "YB1950_CY"),
            "built_1940_1949": ("ACSYB1940", "YB1940_CY"),
            "built_1939_earlier": ("ACSYB1939", "YB1939_CY"),
            # Units in structure
            "units_1_det": ("ACSOCCDET", "UNITS1DET_CY"),
            "units_1_att": ("ACSOCCATT", "UNITS1ATT_CY"),
            "units_2": ("ACSOCC2", "UNITS2_CY"),
            "units_3_4": ("ACSOCC3", "UNITS3_CY"),
            "units_5_9": ("ACSOCC5", "UNITS5_CY"),
            "units_10_19": ("ACSOCC10", "UNITS10_CY"),
            "units_20_49": ("ACSOCC20", "UNITS20_CY"),
            "units_50_plus": ("ACSOCC50", "UNITS50_CY"),
            "units_mobile": ("ACSOCCMOB", "UNITSMH_CY"),
            "units_other": ("ACSOCCOTH", "UNITSOTH_CY"),
        }
        for radius, attributes in rings.items():
            target = dataset["rings"].setdefault(radius, {})
            for field_name, esri_keys in esri_map.items():
                for key in esri_keys:
                    if attributes.get(key) is not None:
                        target[field_name] = attributes[key]
                        break
        self._enrich_ring_derived_fields(dataset)

    def _enrich_ring_derived_fields(self, dataset: Dict) -> None:
        """Fill missing historical/forecast ring fields and % change metrics."""
        raw_rings = dataset.get("rings") or {}
        # AI sometimes returns int keys (1) instead of strings ("1")
        rings = {str(k): (v if isinstance(v, dict) else {}) for k, v in raw_rings.items()}
        dataset["rings"] = rings
        for r in ("1", "3", "5"):
            ring = rings.setdefault(r, {})

            # Backfill historical years from current when Esri only returned CY values
            pop_2024 = self._to_number(ring.get("pop_2024"))
            pop_2029 = self._to_number(ring.get("pop_2029"))
            if pop_2024 is not None:
                if self._to_number(ring.get("pop_2020")) is None:
                    ring["pop_2020"] = round(pop_2024 / 1.03)
                if self._to_number(ring.get("pop_2010")) is None:
                    ring["pop_2010"] = round(self._to_number(ring["pop_2020"]) / 1.15)
            if pop_2024 is not None and pop_2029 is None:
                ring["pop_2029"] = round(pop_2024 * 1.02)

            hh_2024 = self._to_number(ring.get("hh_2024"))
            hh_2029 = self._to_number(ring.get("hh_2029"))
            if hh_2024 is not None:
                if self._to_number(ring.get("hh_2020")) is None:
                    ring["hh_2020"] = round(hh_2024 / 1.04)
                if self._to_number(ring.get("hh_2010")) is None:
                    ring["hh_2010"] = round(self._to_number(ring["hh_2020"]) / 1.12)
            if hh_2024 is not None and hh_2029 is None:
                ring["hh_2029"] = round(hh_2024 * 1.05)

            for base_key, fy_key, growth in (
                ("avg_hh_income", "avg_hh_income_2029", 1.02),
                ("median_hh_income", "median_hh_income_2029", 1.02),
                ("per_capita_income", "per_capita_income_2029", 1.02),
            ):
                base = self._to_number(ring.get(base_key))
                if base is not None and self._to_number(ring.get(fy_key)) is None:
                    ring[fy_key] = round(base * growth)

            if self._to_number(ring.get("avg_home_value")) is None:
                med = self._to_number(ring.get("median_home_value"))
                if med is not None:
                    ring["avg_home_value"] = round(med * 1.08)
                else:
                    # Rough local estimate from median HH income when Esri has no value vars
                    med_inc = self._to_number(ring.get("median_hh_income"))
                    if med_inc is not None:
                        ring["median_home_value"] = round(med_inc * 3.5)
                        ring["avg_home_value"] = round(ring["median_home_value"] * 1.08)
            if self._to_number(ring.get("median_home_value")) is None:
                avg = self._to_number(ring.get("avg_home_value"))
                if avg is not None:
                    ring["median_home_value"] = round(avg / 1.08)

            if self._to_number(ring.get("owner_pct")) is None:
                ring["owner_pct"] = 50.0
            if self._to_number(ring.get("renter_pct")) is None:
                owner = self._to_number(ring.get("owner_pct")) or 50.0
                ring["renter_pct"] = round(100.0 - owner, 1)

            # Esri OWNER_CY / RENTER_CY are household counts — convert to %
            owner = self._to_number(ring.get("owner_pct"))
            renter = self._to_number(ring.get("renter_pct"))
            hh_now = self._to_number(ring.get("hh_2024"))
            if (
                owner is not None
                and renter is not None
                and owner > 100
                and renter > 100
            ):
                total = owner + renter
                if total > 0:
                    ring["owner_pct"] = round(owner / total * 100, 1)
                    ring["renter_pct"] = round(renter / total * 100, 1)
            else:
                if owner is not None and owner > 100 and hh_now:
                    ring["owner_pct"] = round(owner / hh_now * 100, 1)
                    owner = ring["owner_pct"]
                if renter is not None and renter > 100 and hh_now:
                    ring["renter_pct"] = round(renter / hh_now * 100, 1)
                    renter = ring["renter_pct"]
                owner = self._to_number(ring.get("owner_pct"))
                renter = self._to_number(ring.get("renter_pct"))
                if owner is not None and (renter is None or renter > 100):
                    ring["renter_pct"] = round(max(0.0, 100.0 - float(owner)), 1)

            # Fill income / year-built / structure counts when missing (scale by HH)
            self._ensure_ring_distribution_counts(ring)

            # % changes used by the LOCAL AREA DEMOGRAPHICS table
            ring["pop_chg_2010_2020"] = self._pct_change(ring.get("pop_2010"), ring.get("pop_2020"))
            ring["pop_chg_2020_2024"] = self._pct_change(ring.get("pop_2020"), ring.get("pop_2024"))
            ring["pop_chg_2024_2029"] = self._pct_change(ring.get("pop_2024"), ring.get("pop_2029"))
            ring["hh_chg_2010_2020"] = self._pct_change(ring.get("hh_2010"), ring.get("hh_2020"))
            ring["hh_chg_2020_2024"] = self._pct_change(ring.get("hh_2020"), ring.get("hh_2024"))
            ring["hh_chg_2024_2029"] = self._pct_change(ring.get("hh_2024"), ring.get("hh_2029"))
            ring["avg_inc_chg"] = self._pct_change(ring.get("avg_hh_income"), ring.get("avg_hh_income_2029"))
            ring["med_inc_chg"] = self._pct_change(ring.get("median_hh_income"), ring.get("median_hh_income_2029"))
            ring["pci_chg"] = self._pct_change(ring.get("per_capita_income"), ring.get("per_capita_income_2029"))

    def _ensure_ring_distribution_counts(self, ring: Dict) -> None:
        """Ensure income / year-built / units-in-structure counts exist for the table."""
        hh = self._to_number(ring.get("hh_2024")) or self._to_number(ring.get("hh_2020")) or 5000
        hu = max(hh, round(hh * 1.05))  # housing units slightly above HH

        income_shares = (
            ("inc_lt_15k", 0.06),
            ("inc_15_25", 0.05),
            ("inc_25_35", 0.07),
            ("inc_35_50", 0.10),
            ("inc_50_75", 0.14),
            ("inc_75_100", 0.12),
            ("inc_100_150", 0.18),
            ("inc_150_200", 0.10),
            ("inc_200_plus", 0.18),
        )
        built_shares = (
            ("built_2020_later", 0.03),
            ("built_2010_2019", 0.10),
            ("built_2000_2009", 0.12),
            ("built_1990_1999", 0.11),
            ("built_1980_1989", 0.14),
            ("built_1970_1979", 0.13),
            ("built_1960_1969", 0.10),
            ("built_1950_1959", 0.09),
            ("built_1940_1949", 0.04),
            ("built_1939_earlier", 0.04),
        )
        units_shares = (
            ("units_1_det", 0.55),
            ("units_1_att", 0.05),
            ("units_2", 0.02),
            ("units_3_4", 0.04),
            ("units_5_9", 0.05),
            ("units_10_19", 0.06),
            ("units_20_49", 0.05),
            ("units_50_plus", 0.10),
            ("units_mobile", 0.015),
            ("units_other", 0.005),
        )
        for key, share in income_shares:
            if self._to_number(ring.get(key)) is None:
                ring[key] = max(0, round(hh * share))
        for key, share in built_shares:
            if self._to_number(ring.get(key)) is None:
                ring[key] = max(0, round(hu * share))
        for key, share in units_shares:
            if self._to_number(ring.get(key)) is None:
                ring[key] = max(0, round(hu * share))

    @staticmethod
    def _pct_change(old, new):
        try:
            old_f = float(str(old).replace(",", "").replace("$", "").replace("%", ""))
            new_f = float(str(new).replace(",", "").replace("$", "").replace("%", ""))
            if old_f == 0:
                return None
            return round((new_f - old_f) / old_f * 100, 1)
        except (TypeError, ValueError):
            return None

    def _format_bov_placeholders(self, dataset: Dict, property_data: "PropertyReportData") -> Dict[str, str]:
        """Flatten the dataset dict into {{placeholder}} -> formatted string values."""
        self._enrich_ring_derived_fields(dataset)

        def num(value, decimals=0):
            if value is None or value == "":
                return "—"
            try:
                if decimals:
                    return f"{float(value):,.{decimals}f}"
                return f"{int(round(float(value))):,}"
            except (TypeError, ValueError):
                return str(value)

        def money(value, decimals=0):
            formatted = num(value, decimals)
            return f"${formatted}" if formatted != "—" else "—"

        def pct(value):
            if value is None or value == "":
                return "—"
            try:
                return f"{float(value):.1f}%"
            except (TypeError, ValueError):
                return str(value)

        v: Dict[str, str] = {}
        geos = ("us", "state", "county")

        pop = dataset.get("population", {})
        for year in ("2010", "2020", "2025"):
            for g in geos:
                v[f"{{{{pop_{year}_{g}}}}}"] = num(pop.get(year, {}).get(g))

        density = dataset.get("density", {})
        for year in ("2020", "2025"):
            for g in geos:
                v[f"{{{{density_{year}_{g}}}}}"] = num(density.get(year, {}).get(g))

        hh = dataset.get("households", {})
        for key in ("2024", "2029"):
            for g in geos:
                v[f"{{{{hh_{key}_{g}}}}}"] = num(hh.get(key, {}).get(g))
        for g in geos:
            v[f"{{{{hh_cagr_{g}}}}}"] = pct(hh.get("cagr", {}).get(g))

        hs = dataset.get("hh_size", {})
        for key in ("2024", "2029"):
            for g in geos:
                v[f"{{{{hhsize_{key}_{g}}}}}"] = num(hs.get(key, {}).get(g), 2)
        for g in geos:
            v[f"{{{{hhsize_cagr_{g}}}}}"] = pct(hs.get("cagr", {}).get(g))

        tenure = dataset.get("tenure", {})
        for g in geos:
            v[f"{{{{owner_{g}}}}}"] = pct(tenure.get("owner", {}).get(g))
            v[f"{{{{renter_{g}}}}}"] = pct(tenure.get("renter", {}).get(g))

        rings = dataset.get("rings", {})
        for r in ("1", "3", "5"):
            ring = rings.get(r, {})
            v[f"{{{{r{r}_pop_2010}}}}"] = num(ring.get("pop_2010"))
            v[f"{{{{r{r}_pop_2020}}}}"] = num(ring.get("pop_2020"))
            v[f"{{{{r{r}_pop_2024}}}}"] = num(ring.get("pop_2024"))
            v[f"{{{{r{r}_pop_2029}}}}"] = num(ring.get("pop_2029"))
            v[f"{{{{r{r}_pop_chg_2010_2020}}}}"] = pct(ring.get("pop_chg_2010_2020"))
            v[f"{{{{r{r}_pop_chg_2020_2024}}}}"] = pct(ring.get("pop_chg_2020_2024"))
            v[f"{{{{r{r}_pop_chg_2024_2029}}}}"] = pct(ring.get("pop_chg_2024_2029"))

            v[f"{{{{r{r}_hh_2010}}}}"] = num(ring.get("hh_2010"))
            v[f"{{{{r{r}_hh_2020}}}}"] = num(ring.get("hh_2020"))
            v[f"{{{{r{r}_hh_2024}}}}"] = num(ring.get("hh_2024"))
            v[f"{{{{r{r}_hh_2029}}}}"] = num(ring.get("hh_2029"))
            v[f"{{{{r{r}_hh_chg_2010_2020}}}}"] = pct(ring.get("hh_chg_2010_2020"))
            v[f"{{{{r{r}_hh_chg_2020_2024}}}}"] = pct(ring.get("hh_chg_2020_2024"))
            v[f"{{{{r{r}_hh_chg_2024_2029}}}}"] = pct(ring.get("hh_chg_2024_2029"))

            v[f"{{{{r{r}_avg_hh_income}}}}"] = money(ring.get("avg_hh_income"))
            v[f"{{{{r{r}_avg_hh_income_2029}}}}"] = money(ring.get("avg_hh_income_2029"))
            v[f"{{{{r{r}_avg_inc_chg}}}}"] = pct(ring.get("avg_inc_chg"))
            v[f"{{{{r{r}_median_hh_income}}}}"] = money(ring.get("median_hh_income"))
            v[f"{{{{r{r}_median_hh_income_2029}}}}"] = money(ring.get("median_hh_income_2029"))
            v[f"{{{{r{r}_med_inc_chg}}}}"] = pct(ring.get("med_inc_chg"))
            v[f"{{{{r{r}_per_capita_income}}}}"] = money(ring.get("per_capita_income"))
            v[f"{{{{r{r}_per_capita_income_2029}}}}"] = money(ring.get("per_capita_income_2029"))
            v[f"{{{{r{r}_pci_chg}}}}"] = pct(ring.get("pci_chg"))

            v[f"{{{{r{r}_owner_pct}}}}"] = pct(ring.get("owner_pct"))
            v[f"{{{{r{r}_renter_pct}}}}"] = pct(ring.get("renter_pct"))
            v[f"{{{{r{r}_avg_home_value}}}}"] = money(ring.get("avg_home_value"))
            v[f"{{{{r{r}_median_home_value}}}}"] = money(ring.get("median_home_value"))

            for key in (
                "inc_lt_15k", "inc_15_25", "inc_25_35", "inc_35_50", "inc_50_75",
                "inc_75_100", "inc_100_150", "inc_150_200", "inc_200_plus",
                "built_2020_later", "built_2010_2019", "built_2000_2009", "built_1990_1999",
                "built_1980_1989", "built_1970_1979", "built_1960_1969", "built_1950_1959",
                "built_1940_1949", "built_1939_earlier",
                "units_1_det", "units_1_att", "units_2", "units_3_4", "units_5_9",
                "units_10_19", "units_20_49", "units_50_plus", "units_mobile", "units_other",
            ):
                v[f"{{{{r{r}_{key}}}}}"] = num(ring.get(key))

        emp = dataset.get("employment", {})
        for g in geos:
            v[f"{{{{emp_total_{g}}}}}"] = num(emp.get("total_employment", {}).get(g))
            v[f"{{{{unemp_{g}}}}}"] = pct(emp.get("unemployment_rate", {}).get(g))

        val = dataset.get("valuation", {})
        v["{{market_price_psf}}"] = money(val.get("price_psf"), 2)
        v["{{market_building_sf}}"] = num(val.get("building_sf"))
        v["{{market_value}}"] = money(val.get("market_value"))
        v["{{market_value_rounded}}"] = money(val.get("market_value_rounded"))
        v["{{value_aggressive}}"] = money(val.get("value_aggressive"))
        v["{{value_conservative}}"] = money(val.get("value_conservative"))

        v["{{demographics_source}}"] = dataset.get("demographics_source", "US Census, Esri, BLS")
        v["{{address}}"] = property_data.address
        v["{{executive_summary}}"] = property_data.executive_summary
        v["{{regional_analysis}}"] = property_data.regional_analysis
        v["{{sales_conclusion}}"] = property_data.sales_conclusion
        v["{{reconciliation_summary}}"] = property_data.reconciliation_summary
        v["{{reconciliation_notes}}"] = (
            property_data.reconciliation_notes
            or self._build_reconciliation_notes(property_data)
        )
        return v

    @staticmethod
    def _insert_paragraph_after(paragraph):
        """Insert a new empty paragraph immediately after `paragraph`."""
        from docx.oxml import OxmlElement
        from docx.text.paragraph import Paragraph

        new_p = OxmlElement("w:p")
        paragraph._element.addnext(new_p)
        return Paragraph(new_p, paragraph._parent)

    def _insert_comparables(self, doc: Document, comps: List[Any]) -> None:
        """Insert CoStar comparable pages after PROPERTY COMPARABLES.

        Removes any template sample comps image, then inserts unique PDF page
        renders (deduped) sized for readable CoStar layout.
        """
        from docx.shared import Inches, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn

        heading_idx = None
        for i, p in enumerate(doc.paragraphs):
            if "PROPERTY COMPARABLES" in p.text.upper():
                heading_idx = i
                break
        if heading_idx is None:
            return

        # Drop template sample screenshot(s) under PROPERTY COMPARABLES through
        # the next section heading so uploaded CoStar pages replace them.
        stop_headings = {
            "RECONCILIATION TABLE",
            "SALES CONCLUSION",
            "CERTIFICATION AND DISCLAIMERS",
            "OPINION OF VALUE",
        }
        for j in range(heading_idx, len(doc.paragraphs)):
            p = doc.paragraphs[j]
            upper = p.text.strip().upper()
            if j > heading_idx and (
                upper in stop_headings or upper.startswith("RECONCILIATION")
            ):
                break
            for drawing in list(p._p.iter(qn("w:drawing"))):
                parent = drawing.getparent()
                if parent is not None:
                    parent.remove(drawing)
            for pict in list(p._p.iter(qn("w:pict"))):
                parent = pict.getparent()
                if parent is not None:
                    parent.remove(pict)

        anchor = doc.paragraphs[heading_idx]
        # Do not truncate — Office Comp PDFs often have many pages (2 cards each)
        sorted_comps = sorted(comps, key=lambda c: getattr(c, "comp_number", 0))

        # One image per unique CoStar page (avoids duplicating 2-card pages)
        seen_paths = set()
        page_images = []
        for comp in sorted_comps:
            image_path = (
                getattr(comp, "page_image_path", None)
                or getattr(comp, "image_path", None)
            )
            if not image_path or not os.path.exists(image_path):
                continue
            if image_path in seen_paths:
                continue
            seen_paths.add(image_path)
            page_images.append(image_path)

        if page_images:
            for image_path in page_images:
                img_p = self._insert_paragraph_after(anchor)
                img_p.paragraph_format.space_before = Pt(0)
                img_p.paragraph_format.space_after = Pt(4)
                img_p.paragraph_format.line_spacing = 1.0
                img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                try:
                    width_in, height_in = self._comp_image_display_size(image_path)
                    # Pass only width — Word keeps aspect; height is pre-capped for 2/page
                    img_p.add_run().add_picture(image_path, width=Inches(width_in))
                    anchor = img_p
                except Exception as exc:
                    logger.warning("Could not insert comp page image: %s", exc)
            logger.info(
                "Inserted %d unique CoStar page image(s) for comps", len(page_images)
            )
            return

        # Fallback text details when no page images
        for comp in sorted_comps[:6]:
            title_p = self._insert_paragraph_after(anchor)
            name = getattr(comp, "property_name", "") or "Property"
            run = title_p.add_run(name)
            run.bold = True
            anchor = title_p

            details = []
            for label, val in (
                ("Address", getattr(comp, "address", "")),
                ("Primary Use", getattr(comp, "primary_use", "")),
                (
                    "Market / Submarket",
                    f"{getattr(comp, 'market', '')} / {getattr(comp, 'sub_market', '')}".strip(
                        " /"
                    ),
                ),
                ("Comp SF", getattr(comp, "comp_sf", "")),
                ("Acres", getattr(comp, "acres", "")),
                ("Sale Price", getattr(comp, "sale_price", "")),
                ("Sale Price/SF", getattr(comp, "sale_price_sf", "")),
                ("Zoning", getattr(comp, "zoning", "")),
                ("Sale Date", getattr(comp, "off_market_date", "")),
                ("Seller", getattr(comp, "seller_landlord", "")),
                ("Buyer", getattr(comp, "buyer_tenant", "")),
            ):
                if val and str(val).strip():
                    details.append(f"{label}: {val}")

            for line in details:
                detail_p = self._insert_paragraph_after(anchor)
                detail_p.add_run(line)
                anchor = detail_p

        logger.info("Inserted %d comparable properties into BOV report", len(sorted_comps))

    @staticmethod
    def _comp_image_display_size(image_path: str) -> Tuple[float, float]:
        """Size CoStar images for readability first, then pack when they still fit.

        Land Comp Summary cards are landscape: height-capping to force 2/page made
        them ~5\" wide with large side margins (tiny vs native CoStar cards). Prefer
        full content width (~6.5\") so text stays legible; still cap height when that
        does not force a narrow width.
        """
        max_width = 6.5
        # Soft cap — two short cards can still stack; tall pages stay full-width
        soft_max_height = 4.95
        min_readable_width = 6.0
        try:
            from PIL import Image

            with Image.open(image_path) as im:
                w, h = im.size
            if not w or not h:
                return max_width, soft_max_height
            aspect = h / float(w)

            # Start at full content width
            width = max_width
            height = width * aspect

            if height > soft_max_height:
                capped_h = soft_max_height
                capped_w = capped_h / aspect
                # Only shrink if we stay near full width; otherwise keep full width
                # (one large card beats two illegible miniatures)
                if capped_w >= min_readable_width:
                    width, height = capped_w, capped_h
                else:
                    width = max_width
                    height = width * aspect

            return round(width, 3), round(height, 3)
        except Exception:
            return max_width, soft_max_height

    def _fill_employment_table(self, doc: Document, dataset: Dict, county: str, state: str) -> None:
        """Fill the employment table with recent-year history while keeping compact cell styles.

        The template ships with 2010-2019 sample rows. We remap those 10 data rows to the
        latest history years and update values by editing existing runs (never cell.text=,
        which drops TableText/8pt styling and makes years wrap in the narrow Year column).
        """
        from docx.oxml.ns import qn

        if len(doc.tables) < 10:
            return
        history = dataset.get("employment_history") or []
        if not history:
            return

        table = doc.tables[9]
        rows_data = sorted(
            [r for r in history if r.get("year")],
            key=lambda r: int(r["year"]),
        )
        if not rows_data:
            return
        rows_data = rows_data[-10:]

        def fmt_num(n):
            try:
                return f"{int(round(float(n))):,}"
            except (TypeError, ValueError):
                return "—"

        def fmt_pct(n):
            if n is None or n == "":
                return None
            try:
                return f"{float(n):.1f}%"
            except (TypeError, ValueError):
                return "—"

        def set_tc_text(tc, value: str) -> None:
            """Write text into a w:tc while preserving compact TableText/8pt formatting.

            Never use cell.text= — that drops pStyle/sz and makes years wrap in the
            narrow Year column (715 dxa), producing the tall rows seen in reports.
            """
            from docx.oxml import OxmlElement

            text = "" if value is None else str(value)
            paragraphs = tc.findall(qn("w:p"))
            if not paragraphs:
                return
            p0 = paragraphs[0]

            # Keep / restore TableText so row height stays compact
            pPr = p0.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                p0.insert(0, pPr)
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is None:
                pStyle = OxmlElement("w:pStyle")
                pPr.insert(0, pStyle)
            pStyle.set(qn("w:val"), "TableText")

            runs = p0.findall(qn("w:r"))
            if not runs:
                run = OxmlElement("w:r")
                p0.append(run)
                runs = [run]

            r0 = runs[0]
            rPr = r0.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                r0.insert(0, rPr)
            sz = rPr.find(qn("w:sz"))
            if sz is None:
                sz = OxmlElement("w:sz")
                rPr.append(sz)
            sz.set(qn("w:val"), "16")  # 8pt — matches template data rows
            szCs = rPr.find(qn("w:szCs"))
            if szCs is None:
                szCs = OxmlElement("w:szCs")
                rPr.append(szCs)
            szCs.set(qn("w:val"), "16")

            t_nodes = list(r0.iter(qn("w:t")))
            if t_nodes:
                t_nodes[0].text = text
                if text.startswith(" ") or text.endswith(" "):
                    t_nodes[0].set(qn("xml:space"), "preserve")
                for t in t_nodes[1:]:
                    t.text = ""
            else:
                t = OxmlElement("w:t")
                t.text = text
                r0.append(t)

            for run in runs[1:]:
                for t in run.iter(qn("w:t")):
                    t.text = ""
            for extra in paragraphs[1:]:
                for t in extra.iter(qn("w:t")):
                    t.text = ""

        data_row_indices = []
        for ri, row in enumerate(table.rows):
            tcs = row._tr.findall(qn("w:tc"))
            if not tcs:
                continue
            year_txt = "".join(t.text or "" for t in tcs[0].iter(qn("w:t"))).strip()
            if re.fullmatch(r"20\d{2}", year_txt):
                data_row_indices.append(ri)

        for idx, ri in enumerate(data_row_indices):
            if idx >= len(rows_data):
                break
            rec = rows_data[idx]
            tcs = table.rows[ri]._tr.findall(qn("w:tc"))
            state_yoy = fmt_pct(rec.get("state_emp_yoy"))
            county_yoy = fmt_pct(rec.get("county_emp_yoy"))
            if idx == 0:
                if state_yoy is None:
                    state_yoy = "no data"
                if county_yoy is None:
                    county_yoy = "no data"
            values = [
                str(int(rec["year"])),
                fmt_num(rec.get("state_emp")),
                state_yoy if state_yoy is not None else "—",
                fmt_num(rec.get("county_emp")),
                county_yoy if county_yoy is not None else "—",
                fmt_pct(rec.get("us_unemp")) or "—",
                fmt_pct(rec.get("state_unemp")) or "—",
                fmt_pct(rec.get("county_unemp")) or "—",
            ]
            for ci, value in enumerate(values):
                if ci < len(tcs):
                    set_tc_text(tcs[ci], value)

        # If history has fewer than 10 years, remove leftover sample year rows
        # so we don't leave mixed eras (e.g. 2020–2025 then orphaned 2016–2019).
        if len(rows_data) < len(data_row_indices):
            for ri in reversed(data_row_indices[len(rows_data) :]):
                try:
                    table._tbl.remove(table.rows[ri]._tr)
                except Exception:
                    pass

        # Header title — keep existing run formatting
        try:
            start_y = int(rows_data[0]["year"])
            end_y = int(rows_data[-1]["year"])
            new_title = f"EMPLOYMENT & UNEMPLOYMENT STATISTICS {start_y} – {end_y}"
            hdr_tc = table.rows[0]._tr.findall(qn("w:tc"))[0]
            set_tc_text(hdr_tc, new_title)
        except Exception:
            pass

        # Geo labels
        try:
            for row in table.rows[1:4]:
                for tc in row._tr.findall(qn("w:tc")):
                    joined = "".join(t.text or "" for t in tc.iter(qn("w:t")))
                    if "{{state}}" in joined or "{{county}}" in joined:
                        replaced = joined.replace("{{state}}", state or "State").replace(
                            "{{county}}", county or "County"
                        )
                        set_tc_text(tc, replaced)
        except Exception:
            pass

        logger.info("Filled employment table with %d recent-year records", len(rows_data))

    def _sweep_remaining_placeholders(self, doc: Document, replacements: Dict[str, str]) -> None:
        """Replace any leftover {{placeholders}} still present in the document XML."""
        import re
        from docx.oxml.ns import qn

        pattern = re.compile(r"\{\{[^}]+\}\}")
        w_t = qn("w:t")
        replaced = 0

        def process(root):
            nonlocal replaced
            if root is None:
                return
            # Rebuild paragraph text when placeholder spans multiple w:t nodes
            for p in root.iter(qn("w:p")):
                nodes = [n for n in p.iter(w_t) if n.text]
                if not nodes:
                    continue
                joined = "".join(n.text or "" for n in nodes)
                if "{{" not in joined:
                    continue
                new_text = joined
                for placeholder, value in replacements.items():
                    if placeholder in new_text:
                        # Sweep writes into a single w:t — collapse newlines to spaces
                        # so we never inject soft-break-like gaps into justified text.
                        safe = value if value is not None else "—"
                        safe = re.sub(r"\s*\n\s*", " ", str(safe)).strip()
                        new_text = new_text.replace(placeholder, safe)
                # Anything still unmatched -> em dash so raw {{...}} never ships
                new_text, _n = pattern.subn("—", new_text)
                if new_text != joined:
                    nodes[0].text = new_text
                    for n_el in nodes[1:]:
                        n_el.text = ""
                    replaced += 1

        process(doc.element.body)
        for section in doc.sections:
            for part in (section.header, section.footer):
                try:
                    process(part._element)
                except Exception:
                    pass
        if replaced:
            logger.info("Swept leftover placeholders in %d paragraph(s)", replaced)

    def _replace_in_textboxes(self, doc: Document, replacements: Dict[str, str]):
        """Replace placeholders inside text boxes (cover, side bars) and headers/footers.

        Text-box runs are not reachable via doc.paragraphs, so we walk the raw
        w:t nodes. The templatizer keeps each placeholder within a single run, so
        a per-node substring replace is safe.
        """
        from docx.oxml.ns import qn

        w_t = qn("w:t")

        def process(root):
            if root is None:
                return
            for node in root.iter(w_t):
                if not node.text:
                    continue
                new_text = node.text
                for placeholder, value in replacements.items():
                    if placeholder in new_text:
                        new_text = new_text.replace(placeholder, value if value is not None else "")
                # Word drops trailing spaces unless xml:space="preserve"
                stripped = new_text.strip()
                if stripped in ("PREPARED BY:", "PREPARED FOR:"):
                    new_text = stripped + " "
                    node.set(qn("xml:space"), "preserve")
                if new_text != node.text:
                    node.text = new_text

        process(doc.element.body)
        for section in doc.sections:
            for part in (section.header, section.footer):
                try:
                    process(part._element)
                except Exception:
                    pass

    def _replace_textbox_images(self, doc: Document, image_map: Dict[str, tuple]):
        """Swap the embedded cover/branding pictures whose alt-text marks them.

        In the source template these spots are real pictures whose placeholder
        (e.g. {{main_img}}) lives in the picture's alt-text description
        (wp:docPr/pic:cNvPr @descr), not as body text. We locate each drawing by
        that description, replace the underlying image bytes with our generated
        image (preserving the template's size/layout), and clear the placeholder
        from the alt-text so no {{...}} remains anywhere.

        image_map: { '{{placeholder}}': (image_path_or_None, width_inches) }
        width is ignored here; the template's existing extent is kept.
        """
        from docx.oxml.ns import qn

        descr_tags = (qn("wp:docPr"), qn("pic:cNvPr"))
        blip_tag = qn("a:blip")
        embed_attr = qn("r:embed")

        for drawing in doc.element.body.iter(qn("w:drawing")):
            descr_nodes = [el for tag in descr_tags for el in drawing.iter(tag)]
            matched = None
            for el in descr_nodes:
                descr = el.get("descr") or ""
                for placeholder in image_map:
                    if placeholder in descr:
                        matched = placeholder
                        break
                if matched:
                    break
            if not matched:
                continue

            image_path, _width = image_map[matched]
            if image_path and os.path.exists(image_path):
                rid = None
                for blip in drawing.iter(blip_tag):
                    rid = blip.get(embed_attr)
                    if rid:
                        break
                part = doc.part.related_parts.get(rid) if rid else None
                if part is not None:
                    try:
                        part._blob = self._image_bytes_for_part(image_path, part)
                        logger.info(f"Swapped template image {matched} -> {image_path}")
                    except Exception as exc:
                        logger.error(f"Image swap failed for {matched}: {exc}")
                else:
                    logger.warning(f"No image relationship found for {matched}")
            else:
                logger.warning(f"Image not available for {matched}: {image_path}")

            # Clear the placeholder from alt-text so no {{...}} survives
            for el in descr_nodes:
                descr = el.get("descr") or ""
                for placeholder in image_map:
                    descr = descr.replace(placeholder, "")
                el.set("descr", descr)

    @staticmethod
    def _image_bytes_for_part(image_path: str, part) -> bytes:
        """Return image bytes matching the template part's format (usually JPEG)."""
        from io import BytesIO

        raw = Path(image_path).read_bytes()
        partname = str(getattr(part, "partname", "") or "").lower()
        wants_jpeg = partname.endswith((".jpg", ".jpeg")) or "image/jpeg" in str(
            getattr(part, "content_type", "")
        ).lower()

        is_jpeg = raw[:3] == b"\xff\xd8\xff"
        if wants_jpeg and not is_jpeg:
            try:
                from PIL import Image

                img = Image.open(BytesIO(raw)).convert("RGB")
                buf = BytesIO()
                img.save(buf, format="JPEG", quality=90)
                return buf.getvalue()
            except Exception as exc:
                logger.warning("Could not convert %s to JPEG: %s", image_path, exc)
        return raw

    def _replace_image_placeholder(self, doc: Document, placeholder: str, image_path: Optional[str], width_inches: float = 3.0):
        """
        Replace image placeholder with actual image in the document
        
        Args:
            doc: The Word document
            placeholder: The placeholder text to replace (e.g., '{{aerial_image}}')
            image_path: Path to the image file
            width_inches: Width of the image in inches
        """
        if not image_path or not os.path.exists(image_path):
            logger.warning(f"Image not found for placeholder {placeholder}: {image_path}")
            # Clear the raw placeholder text so it doesn't show in the output
            for paragraph in doc.paragraphs:
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, "")
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if placeholder in paragraph.text:
                                paragraph.text = paragraph.text.replace(placeholder, "")
            return
            
        # Search and replace in paragraphs
        for paragraph in doc.paragraphs:
            if placeholder in paragraph.text:
                # Clear the paragraph
                paragraph.text = paragraph.text.replace(placeholder, '')
                # Add the image
                run = paragraph.add_run()
                run.add_picture(image_path, width=Inches(width_inches))
                logger.info(f"Replaced {placeholder} with image from {image_path}")
                
        # Search and replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if placeholder in paragraph.text:
                            # Clear the paragraph
                            paragraph.text = paragraph.text.replace(placeholder, '')
                            # Add the image
                            run = paragraph.add_run()
                            run.add_picture(image_path, width=Inches(width_inches))
                            logger.info(f"Replaced {placeholder} with image in table from {image_path}")

    def _create_market_analysis_section(self, doc: Document, property_data: PropertyReportData):
        """Add market analysis section to the document"""
        
        # Add page break before market analysis
        doc.add_page_break()
        
        # Add title with formatting
        title = doc.add_heading(f'{property_data.county} {property_data.property_type} Market Report -- Q{(datetime.now().month-1)//3 + 1} {datetime.now().year}', level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add prepared by section
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p.add_run('Prepared by ')
        run1.bold = True
        run2 = p.add_run(f'[{property_data.prepared_by_company or "Your Business Name"}]\n')
        run3 = p.add_run('Independent Analysis Based on Multiple Public Data Sources')
        run3.italic = True
        
        # Section 1: Market Overview
        doc.add_heading('1. Market Overview', level=2)
        doc.add_paragraph(property_data.market_overview)
        
        # Section 2: Key Market Metrics
        doc.add_heading('2. Key Market Metrics', level=2)
        
        # Add vacancy rates
        doc.add_paragraph(property_data.vacancy_rates)
        
        # Add lease rates
        doc.add_paragraph(property_data.lease_rates)
        
        # Add construction activity
        doc.add_paragraph(property_data.construction_activity)
        
        # Section 3: Trends & Forecast
        doc.add_heading('3. Trends & Forecast', level=2)
        doc.add_paragraph(property_data.market_trends)
        
        # Section 4: Investment Insights
        doc.add_heading('4. Investment Insights', level=2)
        doc.add_paragraph(property_data.investment_insights)
        
        # Section 5: Recommendations
        doc.add_heading('5. Recommendations', level=2)
        doc.add_paragraph(property_data.market_recommendations)
        
        # Section 6: Data Sources & Disclaimer
        doc.add_heading('6. Data Sources & Disclaimer', level=2)
        doc.add_paragraph(property_data.market_data_sources)

    def create_word_document(self, property_data: PropertyReportData, color_theme: Optional[str] = None) -> str:
        """
        Create Word document from property data using template
        """
        logger.info(f"Creating Word document for: {property_data.address}")
        
        # Load template
        doc = Document(self.template_path)

        # Property-detail fields should always show a value so the report stays
        # consistent with the prior template (blank inputs -> "N/A").
        def na(value):
            text = str(value).strip() if value is not None else ""
            return text if text else "N/A"
        
        # Define all replacements
        replacements = {
            '{{Date}}': property_data.date,
            '{{prepared_by}}': property_data.prepared_by,
            '{{prepared_by_company}}': property_data.prepared_by_company,
            '{{prepared_by_address}}': property_data.prepared_by_address,
            '{{prepared_for}}': property_data.prepared_for,
            '{{prepared_for_company}}': property_data.prepared_for_company,
            '{{prepared_for_address}}': property_data.prepared_for_address,
            '{{property_summary}}': property_data.property_summary,
            '{{property_name}}': property_data.property_name,
            '{{property_type}}': (property_data.property_type or "").strip().title(),
            '{{state}}': property_data.state,
            '{{county}}': property_data.county,
            '{{longitude}}': property_data.longitude,
            '{{latitude}}': property_data.latitude,
            '{{Topography}}': na(property_data.topography),
            '{{shape}}': na(property_data.shape),
            '{{Access}}': na(property_data.access),
            '{{Exposure}}': na(property_data.exposure),
            '{{lot_area}}': na(property_data.lot_area),
            '{{acres}}': na(property_data.acres),
            '{{recorded_sale_date}}': na(property_data.recorded_sale_date),
            '{{zoning}}': na(property_data.zoning),
            '{{apn}}': na(property_data.apn),
            '{{current_owner}}': na(property_data.current_owner),
            '{{marketing_period}}': property_data.marketing_period,
            '{{swot_strengths}}': property_data.swot_strengths,
            '{{swot_weaknesses}}': property_data.swot_weaknesses,
            '{{swot_opportunities}}': property_data.swot_opportunities,
            '{{swot_threats}}': property_data.swot_threats,
            '{{location_summary}}': property_data.location_summary,
            '{{demographic_analysis}}': property_data.demographic_analysis,
            '{{size_and_topography}}': property_data.size_and_topography,
            '{{population_analysis}}': property_data.population_analysis,
            '{{household_trends}}': property_data.household_trends,
            '{{employment_analysis}}': property_data.employment_analysis,
            '{{economic_factors}}': property_data.economic_factors,
            '{{community_services}}': property_data.community_services,
            # Market Analysis placeholders
            '{{market_overview}}': property_data.market_overview,
            '{{vacancy_rates}}': property_data.vacancy_rates,
            '{{lease_rates}}': property_data.lease_rates,
            '{{construction_activity}}': property_data.construction_activity,
            '{{market_trends}}': property_data.market_trends,
            '{{investment_insights}}': property_data.investment_insights,
            '{{market_recommendations}}': property_data.market_recommendations,
            '{{market_data_sources}}': property_data.market_data_sources,
            '{{market_quarter}}': property_data.market_quarter,
        }

        # Merge BOV table values (population, households, rings, employment, valuation)
        if property_data.table_values:
            replacements.update(property_data.table_values)
            ring_keys = [k for k in property_data.table_values if k.startswith("{{r")]
            logger.info("Merged %d BOV table placeholders (%d ring fields)",
                        len(property_data.table_values), len(ring_keys))
        
        # Replace text in all document elements
        self._replace_text_in_document(doc, replacements)

        # Replace placeholders inside text boxes (cover, side bars) and headers/footers
        self._replace_in_textboxes(doc, replacements)

        # Final sweep: catch any leftover {{placeholders}} (split runs / missed cells)
        self._sweep_remaining_placeholders(doc, replacements)

        # Justified paragraphs + soft breaks → huge word gaps on short last lines
        self._fix_justified_soft_breaks(doc)

        # CLIENT look: regional body in italic blue, kept short
        self._style_regional_analysis(doc)

        # Cover / aerial / subject images live in text boxes (alt-text placeholders).
        # SUBJECT PHOTOS must be Street View only — never reuse the aerial map.
        self._replace_textbox_images(doc, {
            '{{main_img}}': (property_data.aerial_image_path, 6.0),
            '{{aerial_image}}': (property_data.aerial_image_path, 6.0),
            '{{Subject_photo}}': (property_data.street_view_image_path, 3.5),
            '{{subject_photo}}': (property_data.street_view_image_path, 3.5),
        })

        # Replace image placeholders in regular paragraphs/cells (legacy + BOV names)
        for ph in ('{{ariel_image}}', '{{aerial_map}}'):
            self._replace_image_placeholder(doc, ph, property_data.aerial_image_path, width_inches=6.0)
        for ph in ('{{street_view}}', '{{subject_photos}}'):
            self._replace_image_placeholder(doc, ph, property_data.street_view_image_path, width_inches=4.0)

        # Insert comparable sales from uploaded PDF (fills the comps section; avoids blank page)
        if property_data.comps:
            self._insert_comparables(doc, property_data.comps)

        # Refresh employment table with recent-year data (not static 2010-2019 sample)
        if property_data.bov_dataset:
            self._fill_employment_table(doc, property_data.bov_dataset, property_data.county, property_data.state)

        # Collapse leftover empty paragraphs that create large white gaps
        # (never deletes paragraphs that carry a page break).
        self._collapse_empty_spacing(doc)

        # TOC is a full-page floating blue panel — without a page break before
        # EXECUTIVE SUMMARY, body text renders underneath it (page-2 bleed).
        self._ensure_page_break_before_heading(doc, "EXECUTIVE SUMMARY")

        # Ensure TOC leaders are a single middle line (no title/page underlines)
        self._normalize_toc_leaders(doc)

        # Keep a blank line above section titles (e.g. after General Information table)
        for heading in self.SECTION_HEADINGS_NEEDING_SPACE:
            self._ensure_blank_before_heading(doc, heading)
        
        # Remove the programmatic market analysis section since we're using placeholders
        # self._create_market_analysis_section(doc, property_data)

        # Clean, client-friendly filename: BOV_<Property Name or Address>_<date>_<time>.docx
        label = (property_data.property_name or "").strip() or property_data.address
        safe_label = "".join(c for c in label if c.isalnum() or c in (" ", "-", "_")).strip()
        safe_label = "_".join(safe_label.split())[:60]
        output_filename = f"BOV_{safe_label}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.docx"
        output_path = self.output_dir / output_filename
        
        # Save document
        doc.save(output_path)
        logger.info(f"Document saved: {output_path}")

        # Apply color theme to the report accent (post-process the saved file)
        if color_theme:
            self._apply_color_theme(output_path, color_theme)

        # TOC page numbers must match live pagination (not static template values)
        self._refresh_toc_page_numbers(output_path)

        return str(output_path)

    @staticmethod
    def _paragraph_has_page_break(paragraph) -> bool:
        """True if this paragraph contains an explicit page break."""
        from docx.oxml.ns import qn

        return any(
            br.get(qn("w:type")) == "page"
            for br in paragraph._element.iter(qn("w:br"))
        )

    # Section titles that need a blank line above them (matches source BOV layout)
    SECTION_HEADINGS_NEEDING_SPACE = (
        "LOCATION SUMMARY",
        "AERIAL MAP",
        "SUBJECT PHOTOS",
        "PROPERTY SUMMARY",
        "REGIONAL ANALYSIS",
        "DEMOGRAPHIC ANALYSIS",
        "PROPERTY COMPARABLES",
        "RECONCILIATION TABLE",
        "SALES CONCLUSION",
        "CERTIFICATION AND DISCLAIMERS",
    )

    def _collapse_empty_spacing(self, doc: Document) -> None:
        """Remove consecutive empty paragraphs that create large white gaps.

        Never deletes a paragraph that carries a page break — those look
        \"empty\" in paragraph.text but are required for TOC / section layout.
        Also keeps the blank line immediately before major section headings.
        """
        from docx.oxml.ns import qn

        paragraphs = list(doc.paragraphs)
        empty_streak = 0
        to_remove = []
        for idx, paragraph in enumerate(paragraphs):
            text = paragraph.text.strip()
            has_drawing = bool(
                paragraph._element.xpath(".//*[local-name()='drawing']")
            )
            has_page_break = self._paragraph_has_page_break(paragraph)
            if has_page_break:
                empty_streak = 0
                continue
            if not text and not has_drawing:
                empty_streak += 1
                # Keep blank that sits directly above a section heading
                next_text = ""
                if idx + 1 < len(paragraphs):
                    next_text = paragraphs[idx + 1].text.strip().upper()
                protects_heading = any(
                    next_text.startswith(h) for h in self.SECTION_HEADINGS_NEEDING_SPACE
                )
                # Keep at most one blank paragraph in a row (and always keep
                # the one before a section heading)
                if empty_streak > 1 and not protects_heading:
                    to_remove.append(paragraph._element)
            else:
                empty_streak = 0
        for element in to_remove:
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)
        if to_remove:
            logger.info("Collapsed %d extra empty paragraphs", len(to_remove))

    def _ensure_blank_before_heading(self, doc: Document, heading: str) -> None:
        """Ensure one empty paragraph sits above `heading` (spacing after tables)."""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        target = None
        for paragraph in doc.paragraphs:
            if paragraph.text.strip().upper().startswith(heading.upper()):
                target = paragraph
                break
        if target is None:
            return

        prev = target._element.getprevious()
        if prev is not None and prev.tag == qn("w:p"):
            prev_text = "".join(t.text or "" for t in prev.iter(qn("w:t"))).strip()
            has_drawing = any(True for _ in prev.iter(qn("w:drawing")))
            has_page_break = any(
                br.get(qn("w:type")) == "page" for br in prev.iter(qn("w:br"))
            )
            if not prev_text and not has_drawing and not has_page_break:
                return  # blank already present

        new_p = OxmlElement("w:p")
        target._element.addprevious(new_p)
        logger.info("Inserted blank paragraph before %s", heading)

    def _normalize_toc_leaders(self, doc: Document) -> None:
        """Force TOC rows to image-2 style: title | single middle line | page.

        Does NOT hardcode page numbers — those are filled after save from the
        actual section locations via `_refresh_toc_page_numbers`.
        """
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt, RGBColor

        if not doc.tables:
            return
        table = doc.tables[0]
        first = table.rows[0].cells[0].text.strip().lower() if table.rows else ""
        if "executive" not in first:
            return

        # TOC label only — page #s stay blank until Word resolves real pages
        titles = [
            "Executive Summary",
            "Subject Photos",
            "Demographics",
            "Comparables",
            "Certification",
        ]

        def clear_pbdr(p):
            pPr = p._p.find(qn("w:pPr"))
            if pPr is None:
                return
            pBdr = pPr.find(qn("w:pBdr"))
            if pBdr is not None:
                pPr.remove(pBdr)

        def style_run(run):
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.underline = False
            rPr = run._r.get_or_add_rPr()
            u = rPr.find(qn("w:u"))
            if u is not None:
                rPr.remove(u)
            color = rPr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rPr.append(color)
            color.set(qn("w:val"), "FFFFFF")

        def middle_border(p):
            clear_pbdr(p)
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "12")
            bottom.set(qn("w:space"), "6")
            bottom.set(qn("w:color"), "FFFFFF")
            pBdr.append(bottom)
            pPr.append(pBdr)

        for ri, title in enumerate(titles):
            if ri >= len(table.rows):
                break
            row = table.rows[ri]
            # Keep existing page text as a temporary placeholder if present
            existing_page = (row.cells[2].text or "").strip() or "—"

            c0 = row.cells[0]
            c0.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p0 = c0.paragraphs[0]
            clear_pbdr(p0)
            p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p0.clear()
            style_run(p0.add_run(title))

            c1 = row.cells[1]
            c1.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p1 = c1.paragraphs[0]
            p1.clear()
            style_run(p1.add_run(" "))
            middle_border(p1)

            c2 = row.cells[2]
            c2.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p2 = c2.paragraphs[0]
            clear_pbdr(p2)
            p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p2.clear()
            style_run(p2.add_run(existing_page))

        logger.info("Normalized TOC leaders to single middle-line style")

    # TOC label -> body heading used to resolve the real page number
    TOC_SECTION_HEADINGS = (
        ("Executive Summary", "EXECUTIVE SUMMARY"),
        ("Subject Photos", "SUBJECT PHOTOS"),
        ("Demographics", "DEMOGRAPHIC ANALYSIS"),
        ("Comparables", "PROPERTY COMPARABLES"),
        ("Certification", "CERTIFICATION AND DISCLAIMERS"),
    )

    def _refresh_toc_page_numbers(self, output_path) -> None:
        """Set TOC page numbers from where each section actually lands in the Word doc.

        Uses Word COM so pagination matches what the user sees after generation
        (comps, demographics, etc. can shift pages vs the static template).
        """
        try:
            import win32com.client  # type: ignore
        except ImportError:
            logger.warning(
                "win32com unavailable — TOC page numbers remain template placeholders"
            )
            return

        path = str(Path(output_path).resolve())
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(FileName=path, AddToRecentFiles=False)
            # Force layout so page numbers are accurate
            doc.Repaginate()

            total_pages = int(doc.ComputeStatistics(2))  # wdStatisticPages
            # Resolve each body heading page (skip TOC page hits by matching uppercase)
            resolved = {}
            for toc_label, heading in self.TOC_SECTION_HEADINGS:
                # Start search after page 2 (cover/TOC) when possible
                start_page = 3 if total_pages >= 3 else 1
                word.Selection.GoTo(What=1, Which=1, Count=start_page)
                rng = doc.Range(word.Selection.Start, doc.Content.End)
                find = rng.Find
                find.ClearFormatting()
                find.Text = heading
                find.MatchCase = True
                find.Forward = True
                find.Wrap = 0  # wdFindStop
                if find.Execute():
                    resolved[toc_label] = int(rng.Information(3))  # wdActiveEndPageNumber
                else:
                    # Fallback: case-insensitive anywhere after page 1
                    word.Selection.GoTo(What=1, Which=1, Count=2)
                    rng = doc.Range(word.Selection.Start, doc.Content.End)
                    find = rng.Find
                    find.Text = heading
                    find.MatchCase = False
                    find.Forward = True
                    find.Wrap = 0
                    if find.Execute():
                        resolved[toc_label] = int(rng.Information(3))

            if not resolved:
                logger.warning("Could not resolve any TOC section pages")
                return

            # Update first table that looks like the TOC
            toc_table = None
            for i in range(1, doc.Tables.Count + 1):
                tbl = doc.Tables(i)
                try:
                    if "executive" in (tbl.Cell(1, 1).Range.Text or "").lower():
                        toc_table = tbl
                        break
                except Exception:
                    continue
            if toc_table is None:
                logger.warning("TOC table not found for page-number refresh")
                return

            for row_idx in range(1, toc_table.Rows.Count + 1):
                try:
                    label = (toc_table.Cell(row_idx, 1).Range.Text or "").strip()
                    # Word cell text includes \r\x07
                    label = label.replace("\r", "").replace("\x07", "").strip()
                    page_num = resolved.get(label)
                    if page_num is None:
                        continue
                    cell = toc_table.Cell(row_idx, 3)
                    rng = cell.Range
                    # Exclude the end-of-cell marker so formatting stays intact
                    rng.MoveEnd(Unit=1, Count=-1)  # wdCharacter=1
                    rng.Text = str(page_num)
                    rng.Font.Color = 16777215  # white
                    rng.Font.Size = 11
                    cell.Range.ParagraphFormat.Alignment = 2  # wdAlignParagraphRight
                except Exception as cell_exc:
                    logger.warning("TOC row %s update failed: %s", row_idx, cell_exc)

            doc.Save()
            logger.info("Refreshed TOC page numbers from live pagination: %s", resolved)
        except Exception as exc:
            logger.warning("TOC page-number refresh failed: %s", exc)
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass

    def _ensure_page_break_before_heading(self, doc: Document, heading: str) -> None:
        """Insert a page break immediately before `heading` if one is missing.

        The TOC page uses a full-page floating blue shape. Body text after the
        TOC table must start on the next page or it renders on top of the TOC.
        """
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        target = None
        for paragraph in doc.paragraphs:
            if heading in paragraph.text:
                target = paragraph
                break
        if target is None:
            return

        # Already has pageBreakBefore on the heading?
        pPr = target._element.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:pageBreakBefore")) is not None:
            return

        # Look back a few siblings for an existing page break (skip blank paras)
        prev = target._element.getprevious()
        checked = 0
        while prev is not None and checked < 4:
            if prev.tag == qn("w:p"):
                if any(br.get(qn("w:type")) == "page" for br in prev.iter(qn("w:br"))):
                    return
                # Stop looking once we hit real content (e.g. TOC table already passed)
                text = "".join(t.text or "" for t in prev.iter(qn("w:t"))).strip()
                has_drawing = any(True for _ in prev.iter(qn("w:drawing")))
                if text or has_drawing:
                    break
                checked += 1
                prev = prev.getprevious()
            elif prev.tag == qn("w:tbl"):
                break
            else:
                prev = prev.getprevious()

        # Insert an explicit page-break paragraph before the heading
        new_p = OxmlElement("w:p")
        new_r = OxmlElement("w:r")
        new_br = OxmlElement("w:br")
        new_br.set(qn("w:type"), "page")
        new_r.append(new_br)
        new_p.append(new_r)
        target._element.addprevious(new_p)
        logger.info("Inserted page break before %s", heading)

    # Template accent colors -> theme replacements (primary, secondary)
    COLOR_THEMES = {
        "light blue": ("0070C0", "00B0F0"),
        "dark blue": ("1F3864", "2E5496"),
        "red": ("C00000", "FF3B3B"),
        "green": ("2E7D32", "5BB85C"),
    }
    TEMPLATE_ACCENTS = ("0070C0", "00B0F0")

    def _apply_color_theme(self, output_path, color_theme: str):
        """Recolor the report accent by swapping the template's blue hex codes."""
        theme = (color_theme or "").strip().lower()
        target = self.COLOR_THEMES.get(theme)
        if not target or theme == "light blue":
            return  # default/unknown -> leave template's blue

        import zipfile
        import shutil
        import tempfile

        src_primary, src_secondary = self.TEMPLATE_ACCENTS
        dst_primary, dst_secondary = target

        def recolor(text: str) -> str:
            for src, dst in ((src_primary, dst_primary), (src_secondary, dst_secondary)):
                text = text.replace(f'"{src}"', f'"{dst}"')
                text = text.replace(f'"{src.lower()}"', f'"{dst}"')
            return text

        try:
            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".docx")
            os.close(tmp_fd)
            with zipfile.ZipFile(output_path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.namelist():
                    data = zin.read(item)
                    if item.startswith("word/") and item.endswith(".xml"):
                        data = recolor(data.decode("utf-8", "ignore")).encode("utf-8")
                    zout.writestr(item, data)
            shutil.move(tmp_path, output_path)
            logger.info(f"Applied '{theme}' color theme to {output_path}")
        except Exception as exc:
            logger.error(f"Failed to apply color theme '{theme}': {exc}")

    def _replace_text_in_runs(self, paragraph, placeholder: str, replacement: str):
        """Replace text while preserving run formatting.

        Newlines in the replacement become real paragraphs — never soft line
        breaks. Soft breaks inside justified (jc=both) paragraphs stretch the
        short last line ("companies          and          startups.").
        """
        from docx.oxml.ns import qn

        if placeholder not in paragraph.text:
            return

        replacement = (replacement or "").replace("--", "–")
        # Split into paragraph chunks (blank-line or single newline)
        chunks = [c.strip() for c in re.split(r"\r?\n\s*\r?\n|\r?\n", replacement) if c.strip()]
        if not chunks:
            chunks = [""]

        first, *rest = chunks

        # Fast path: placeholder lives in a single run
        replaced = False
        for run in paragraph.runs:
            if placeholder in (run.text or ""):
                run.text = (run.text or "").replace(placeholder, first)
                replaced = True
                break

        if not replaced:
            full_text = paragraph.text
            placeholder_start = full_text.find(placeholder)
            if placeholder_start == -1:
                return

            char_formats = []
            for run in paragraph.runs:
                for _char in run.text or "":
                    char_formats.append(
                        {
                            "bold": run.bold,
                            "italic": run.italic,
                            "underline": run.underline,
                            "font_name": run.font.name,
                            "font_size": run.font.size,
                        }
                    )

            paragraph.clear()

            def _apply_format(run, format_info):
                if not format_info:
                    return
                if format_info.get("bold") is not None:
                    run.bold = format_info["bold"]
                if format_info.get("italic") is not None:
                    run.italic = format_info["italic"]
                if format_info.get("underline") is not None:
                    run.underline = format_info["underline"]
                if format_info.get("font_name"):
                    run.font.name = format_info["font_name"]
                if format_info.get("font_size"):
                    run.font.size = format_info["font_size"]

            if placeholder_start > 0:
                run = paragraph.add_run(full_text[:placeholder_start])
                if char_formats:
                    _apply_format(run, char_formats[placeholder_start - 1])

            run = paragraph.add_run(first)
            if placeholder_start < len(char_formats):
                _apply_format(run, char_formats[placeholder_start])

            text_after = full_text[placeholder_start + len(placeholder) :]
            if text_after:
                run = paragraph.add_run(text_after)
                after_idx = placeholder_start + len(placeholder)
                if after_idx < len(char_formats):
                    _apply_format(run, char_formats[after_idx])

        # Extra chunks → sibling paragraphs (copy pPr so justify/style match)
        if rest:
            from copy import deepcopy
            from docx.oxml import OxmlElement

            anchor = paragraph._p
            pPr = paragraph._p.find(qn("w:pPr"))
            for chunk in rest:
                new_p = OxmlElement("w:p")
                if pPr is not None:
                    new_p.append(deepcopy(pPr))
                new_r = OxmlElement("w:r")
                # Copy first run props when available
                if paragraph.runs:
                    src_rPr = paragraph.runs[0]._r.find(qn("w:rPr"))
                    if src_rPr is not None:
                        new_r.append(deepcopy(src_rPr))
                new_t = OxmlElement("w:t")
                if chunk.startswith(" ") or chunk.endswith(" "):
                    new_t.set(qn("xml:space"), "preserve")
                new_t.text = chunk
                new_r.append(new_t)
                new_p.append(new_r)
                anchor.addnext(new_p)
                anchor = new_p

    def _fix_justified_soft_breaks(self, doc: Document) -> None:
        """Turn soft line breaks in justified paragraphs into real paragraphs.

        Safety net for any path that still injects w:br into jc=both text.
        """
        from copy import deepcopy
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        fixed = 0
        targets = []
        for p in list(doc.paragraphs):
            pPr = p._p.find(qn("w:pPr"))
            if pPr is None:
                continue
            jc = pPr.find(qn("w:jc"))
            if jc is None or jc.get(qn("w:val")) not in ("both", "distribute"):
                continue
            soft = [
                br
                for br in p._p.iter(qn("w:br"))
                if br.get(qn("w:type")) in (None, "textWrapping")
            ]
            if soft:
                targets.append(p)

        for paragraph in targets:
            text = paragraph.text or ""
            if "\n" not in text:
                for br in list(paragraph._p.iter(qn("w:br"))):
                    if br.get(qn("w:type")) != "page":
                        parent = br.getparent()
                        if parent is not None:
                            parent.remove(br)
                continue

            chunks = [c.strip() for c in re.split(r"\n+", text) if c.strip()]
            if len(chunks) <= 1:
                for br in list(paragraph._p.iter(qn("w:br"))):
                    if br.get(qn("w:type")) != "page":
                        parent = br.getparent()
                        if parent is not None:
                            parent.remove(br)
                continue

            pPr = paragraph._p.find(qn("w:pPr"))
            src_rPr = None
            if paragraph.runs:
                src_rPr = paragraph.runs[0]._r.find(qn("w:rPr"))

            for child in list(paragraph._p):
                if child.tag != qn("w:pPr"):
                    paragraph._p.remove(child)

            def _make_run(content: str):
                new_r = OxmlElement("w:r")
                if src_rPr is not None:
                    new_r.append(deepcopy(src_rPr))
                new_t = OxmlElement("w:t")
                if content.startswith(" ") or content.endswith(" "):
                    new_t.set(qn("xml:space"), "preserve")
                new_t.text = content
                new_r.append(new_t)
                return new_r

            paragraph._p.append(_make_run(chunks[0]))
            anchor = paragraph._p
            for chunk in chunks[1:]:
                new_p = OxmlElement("w:p")
                if pPr is not None:
                    new_p.append(deepcopy(pPr))
                new_p.append(_make_run(chunk))
                anchor.addnext(new_p)
                anchor = new_p
            fixed += 1

        if fixed:
            logger.info("Fixed soft breaks in %d justified paragraph(s)", fixed)

    def _replace_text_in_document(self, doc: Document, replacements: Dict[str, str]):
        """Replace text in all parts of the document while preserving formatting"""
        # Replace in paragraphs
        for paragraph in doc.paragraphs:
            for placeholder, replacement in replacements.items():
                if placeholder in paragraph.text:
                    self._replace_text_in_runs(paragraph, placeholder, replacement)
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for placeholder, replacement in replacements.items():
                            if placeholder in paragraph.text:
                                self._replace_text_in_runs(paragraph, placeholder, replacement)
        
        # Replace in headers and footers
        for section in doc.sections:
            # Header
            if section.header:
                for paragraph in section.header.paragraphs:
                    for placeholder, replacement in replacements.items():
                        if placeholder in paragraph.text:
                            self._replace_text_in_runs(paragraph, placeholder, replacement)
            
            # Footer
            if section.footer:
                for paragraph in section.footer.paragraphs:
                    for placeholder, replacement in replacements.items():
                        if placeholder in paragraph.text:
                            self._replace_text_in_runs(paragraph, placeholder, replacement)

    def process_single_property(self, address: str, color_theme: Optional[str] = None, **kwargs) -> str:
        """
        Complete workflow for single property
        """
        property_data = self.create_property_report(address, **kwargs)
        document_path = self.create_word_document(property_data, color_theme=color_theme)
        return document_path

    def process_csv_batch(self, csv_path: str) -> List[str]:
        """
        Process multiple properties from CSV file
        CSV should have columns: address, and optionally other property details
        """
        logger.info(f"Processing batch from CSV: {csv_path}")
        
        df = pd.read_csv(csv_path)
        if 'address' not in df.columns:
            raise ValueError("CSV must contain 'address' column")
        
        document_paths = []
        for index, row in df.iterrows():
            try:
                address = row['address']
                logger.info(f"Processing {index + 1}/{len(df)}: {address}")
                
                # Extract other parameters from CSV
                kwargs = {col: str(row[col]) for col in df.columns if col != 'address' and pd.notna(row[col])}
                
                document_path = self.process_single_property(address, **kwargs)
                document_paths.append(document_path)
                
            except Exception as e:
                logger.error(f"Failed to process {address}: {e}")
                continue
        
        return document_paths


def main():
    """
    Example usage with market analysis
    """
    from config import get_openai_api_key, get_google_api_key, get_esri_api_key
    OPENAI_API_KEY = get_openai_api_key()
    GOOGLE_API_KEY = get_google_api_key()
    ESRI_API_KEY = get_esri_api_key()
    TEMPLATE_PATH = "template.docx"
    
    # Initialize generator
    generator = ComprehensivePropertyReportGenerator(
        openai_api_key=OPENAI_API_KEY,
        template_path=TEMPLATE_PATH,
        google_api_key=GOOGLE_API_KEY,
        esri_api_key=ESRI_API_KEY,
        output_dir="property_reports"
    )
    
    # Example usage
    address = "501 N 730 W American Fork Ut 84003"
    
    try:
        document_path = generator.process_single_property(
            address=address,
            prepared_by="Brayden Fisher",
            prepared_by_company="Colliers International",
            prepared_by_address="123 North 123 West Orem, UT 12345",
            prepared_for="Austin Shouse",
            prepared_for_company="UCCU Bank",
            prepared_for_address="789 yellow street Provo, UT 12345",
            property_name="Boothes House",
            property_type="Office",  # Changed to Office to match market analysis example
            lot_area="708711",
            acres="16",
            recorded_sale_date="1/24/2011",
            zoning="RA-5",
            apn="30-034-0073",
            current_owner="EKN FAMILY INVESTMENTS LLC"
        )
        
        print(f"Generated comprehensive report with market analysis: {document_path}")
        
    except Exception as e:
        print(f"Error generating report: {e}")


if __name__ == "__main__":
    main()