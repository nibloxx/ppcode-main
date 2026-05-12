# Property Report API Server

This Flask API server integrates both `both4.py` (Property Report Generator) and `comp2.py` (Comparable Properties Extractor) functionality, allowing you to call them from Postman or any HTTP client.

## Setup

1. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Start the Server:**
   ```bash
   python api_server.py
   ```

   The server will run on `http://localhost:5000`

## API Endpoints

### 1. Health Check
- **URL:** `GET /health`
- **Description:** Check if the server is running and generators are initialized
- **Response:**
  ```json
  {
    "status": "healthy",
    "timestamp": "2024-01-15T10:30:00",
    "property_generator": true,
    "comp_extractor": true
  }
  ```

### 2. Generate Property Report
- **URL:** `POST /generate_property_report`
- **Description:** Generate a comprehensive property report using `both4.py`
- **Content-Type:** `application/json`
- **Request Body:**
  ```json
  {
    "address": "501 N 730 W American Fork Ut 84003",
    "prepared_by": "Brayden Fisher",
    "prepared_by_company": "Colliers International",
    "prepared_by_address": "123 North 123 West Orem, UT 12345",
    "prepared_for": "Austin Shouse",
    "prepared_for_company": "UCCU Bank",
    "prepared_for_address": "789 yellow street Provo, UT 12345",
    "property_name": "Boothes House",
    "property_type": "Office",
    "lot_area": "708711",
    "acres": "16",
    "recorded_sale_date": "1/24/2011",
    "zoning": "RA-5",
    "apn": "30-034-0073",
    "current_owner": "EKN FAMILY INVESTMENTS LLC"
  }
  ```
- **Response:**
  ```json
  {
    "success": true,
    "message": "Property report generated successfully",
    "document_path": "property_reports/Property_Report_501 N 730 W American Fork Ut 84003_20240115_103000.docx",
    "filename": "Property_Report_501 N 730 W American Fork Ut 84003_20240115_103000.docx"
  }
  ```

### 3. Download Property Report
- **URL:** `GET /download_property_report/<filename>`
- **Description:** Download a generated property report file
- **Example:** `GET /download_property_report/Property_Report_501 N 730 W American Fork Ut 84003_20240115_103000.docx`

### 4. Process Comparable Properties
- **URL:** `POST /process_comps`
- **Description:** Extract comps from PDF and generate report using `comp2.py`
- **Content-Type:** `multipart/form-data` or `application/json`

#### Option A: File Upload (multipart/form-data)
- **Form Fields:**
  - `pdf_file`: PDF file containing comp data (required)
  - `template_file`: Word template file (optional, uses default if not provided)

#### Option B: File Paths (application/json)
- **Request Body:**
  ```json
  {
    "pdf_path": "comp.pdf",
    "template_path": "comptemplate.docx"
  }
  ```
- **Response:**
  ```json
  {
    "success": true,
    "message": "Processed 6 comparable properties",
    "document_path": "property_reports/Report_with_Comps_20240115_103000.docx",
    "filename": "Report_with_Comps_20240115_103000.docx",
    "comps_count": 6,
    "comps_data": [
      {
        "comp_number": 1,
        "property_name": "Sample Property",
        "address": "123 Main St, City, State 12345",
        "sale_price": "$500,000",
        "sale_price_sf": "$25.00",
        "market": "Utah County",
        "sub_market": "Provo/Orem"
      }
    ]
  }
  ```

### 5. Download Comp Report
- **URL:** `GET /download_comp_report/<filename>`
- **Description:** Download a generated comp report file
- **Example:** `GET /download_comp_report/Report_with_Comps_20240115_103000.docx`

### 6. Generate Combined Report
- **URL:** `POST /generate_combined_report`
- **Description:** Generate both property report and comp report in one call
- **Content-Type:** `application/json`
- **Request Body:**
  ```json
  {
    "address": "501 N 730 W American Fork Ut 84003",
    "property_data": {
      "prepared_by": "Brayden Fisher",
      "prepared_by_company": "Colliers International",
      "property_type": "Office"
    },
    "pdf_path": "comp.pdf",
    "template_path": "comptemplate.docx"
  }
  ```

### 7. List Reports
- **URL:** `GET /list_reports`
- **Description:** List all generated reports
- **Response:**
  ```json
  {
    "reports": [
      {
        "filename": "Property_Report_501 N 730 W American Fork Ut 84003_20240115_103000.docx",
        "size_bytes": 1024000,
        "created": "2024-01-15T10:30:00",
        "modified": "2024-01-15T10:30:00"
      }
    ],
    "total_count": 1
  }
  ```

### 8. Delete Report
- **URL:** `DELETE /delete_report/<filename>`
- **Description:** Delete a specific report file
- **Example:** `DELETE /delete_report/Property_Report_501 N 730 W American Fork Ut 84003_20240115_103000.docx`

## Postman Setup

### 1. Create a New Collection
1. Open Postman
2. Create a new collection called "Property Report API"

### 2. Add Requests

#### Health Check
- **Method:** GET
- **URL:** `http://localhost:5000/health`

#### Generate Property Report
- **Method:** POST
- **URL:** `http://localhost:5000/generate_property_report`
- **Headers:** `Content-Type: application/json`
- **Body (raw JSON):**
  ```json
  {
    "address": "501 N 730 W American Fork Ut 84003",
    "prepared_by": "Brayden Fisher",
    "prepared_by_company": "Colliers International",
    "prepared_by_address": "123 North 123 West Orem, UT 12345",
    "prepared_for": "Austin Shouse",
    "prepared_for_company": "UCCU Bank",
    "prepared_for_address": "789 yellow street Provo, UT 12345",
    "property_name": "Boothes House",
    "property_type": "Office",
    "lot_area": "708711",
    "acres": "16",
    "recorded_sale_date": "1/24/2011",
    "zoning": "RA-5",
    "apn": "30-034-0073",
    "current_owner": "EKN FAMILY INVESTMENTS LLC"
  }
  ```

#### Process Comps (File Upload)
- **Method:** POST
- **URL:** `http://localhost:5000/process_comps`
- **Body (form-data):**
  - Key: `pdf_file`, Type: File, Value: Select your PDF file
  - Key: `template_file`, Type: File, Value: Select your template file (optional)

#### Process Comps (File Paths)
- **Method:** POST
- **URL:** `http://localhost:5000/process_comps`
- **Headers:** `Content-Type: application/json`
- **Body (raw JSON):**
  ```json
  {
    "pdf_path": "comp.pdf",
    "template_path": "comptemplate.docx"
  }
  ```

#### Download Report
- **Method:** GET
- **URL:** `http://localhost:5000/download_property_report/{{filename}}`
- Replace `{{filename}}` with the actual filename from the response

#### List Reports
- **Method:** GET
- **URL:** `http://localhost:5000/list_reports`

### 3. Environment Variables (Optional)
Create an environment in Postman with variables:
- `base_url`: `http://localhost:5000`
- `filename`: (will be set from responses)

## File Structure

```
ppcode/
├── api_server.py          # Flask API server
├── both4.py              # Property report generator
├── comp2.py              # Comp extractor
├── template.docx         # Property report template
├── comptemplate.docx     # Comp report template
├── comp.pdf             # Sample comp PDF
├── requirements.txt      # Python dependencies
├── README_API.md        # This file
└── property_reports/    # Generated reports directory
    ├── images/          # Property images
    └── comp_images/     # Comp images
```

## Error Handling

The API returns appropriate HTTP status codes:
- `200`: Success
- `400`: Bad Request (missing required fields)
- `404`: File not found
- `500`: Internal server error

Error responses include a descriptive message:
```json
{
  "error": "Failed to generate property report: Invalid address"
}
```

## Notes

1. **API Keys:** The server uses hardcoded API keys for OpenAI and Google Maps. In production, these should be environment variables.

2. **File Storage:** All generated files are stored in the `property_reports/` directory.

3. **Templates:** Make sure `template.docx` and `comptemplate.docx` exist in the root directory.

4. **CORS:** The API has CORS enabled for cross-origin requests.

5. **Logging:** All operations are logged for debugging purposes.

## Troubleshooting

1. **Server won't start:** Check if all dependencies are installed
2. **Property generator fails:** Verify API keys and template file existence
3. **Comp extraction fails:** Ensure PDF file is valid and contains comp data
4. **File download fails:** Check if the file exists in the property_reports directory 