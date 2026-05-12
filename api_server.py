from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
import shutil
from pathlib import Path
import logging
from datetime import datetime
import traceback
import requests
import json

# Import the classes from both scripts
from both4 import ComprehensivePropertyReportGenerator
from comp2 import CompExtractor

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Configuration
OPENAI_API_KEY = 'sk-rSNepa2p0yahWHyJ27CuZd2snzGme9R2hlNOpreUWDT3BlbkFJLVmbTqYzzOpIt-K5mexrZw5W37ziLI2jkNiz0NmmAA'
GOOGLE_API_KEY = 'AIzaSyCl6Oc03tJ-MkQEXMc84pF9lXURvPLPmHU'
TEMPLATE_PATH = "template.docx"
COMP_TEMPLATE_PATH = "comptemplate.docx"

# Initialize generators
property_generator = None
comp_extractor = None

def initialize_generators():
    """Initialize the property generator and comp extractor"""
    global property_generator, comp_extractor
    
    try:
        property_generator = ComprehensivePropertyReportGenerator(
            openai_api_key=OPENAI_API_KEY,
            google_api_key=GOOGLE_API_KEY,
            template_path=TEMPLATE_PATH,
            output_dir="property_reports"
        )
        logger.info("Property generator initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize property generator: {e}")
        property_generator = None
    
    try:
        comp_extractor = CompExtractor(output_dir="property_reports")
        logger.info("Comp extractor initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize comp extractor: {e}")
        comp_extractor = None

# Initialize generators when module is imported
initialize_generators()

def download_pdf_from_url(url: str, temp_dir: Path) -> Path:
    """Download PDF from URL and return local file path"""
    try:
        logger.info(f"Downloading PDF from URL: {url}")
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        pdf_path = temp_dir / "downloaded_comp.pdf"
        with open(pdf_path, 'wb') as f:
            f.write(response.content)
        
        logger.info(f"PDF downloaded successfully to: {pdf_path}")
        return pdf_path
        
    except Exception as e:
        logger.error(f"Error downloading PDF from URL: {e}")
        raise Exception(f"Failed to download PDF from URL: {str(e)}")

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'property_generator': property_generator is not None,
        'comp_extractor': comp_extractor is not None
    })

@app.route('/test_upload', methods=['POST'])
def test_upload():
    """Test endpoint for file uploads"""
    try:
        if not request.files:
            return jsonify({'error': 'No files uploaded'}), 400
        
        files_info = []
        for field_name, file in request.files.items():
            if file.filename:
                files_info.append({
                    'field_name': field_name,
                    'filename': file.filename,
                    'content_type': file.content_type,
                    'size': len(file.read())
                })
                file.seek(0)  # Reset file pointer
        
        form_data = {}
        for key, value in request.form.items():
            form_data[key] = value
        
        return jsonify({
            'success': True,
            'message': 'Upload test successful',
            'files': files_info,
            'form_data': form_data
        })
        
    except Exception as e:
        logger.error(f"Error in test upload: {e}")
        return jsonify({'error': f'Upload test failed: {str(e)}'}), 500

@app.route('/reinitialize', methods=['POST'])
def reinitialize():
    """Reinitialize the generators"""
    try:
        initialize_generators()
        return jsonify({
            'success': True,
            'message': 'Generators reinitialized',
            'property_generator': property_generator is not None,
            'comp_extractor': comp_extractor is not None
        })
    except Exception as e:
        logger.error(f"Error reinitializing generators: {e}")
        return jsonify({
            'error': f'Failed to reinitialize generators: {str(e)}'
        }), 500

@app.route('/generate_property_report', methods=['POST'])
def generate_property_report():
    """
    Generate a comprehensive property report
    
    Expected JSON payload:
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
    """
    if not property_generator:
        return jsonify({'error': 'Property generator not initialized'}), 500
    
    try:
        data = request.get_json()
        
        if not data or 'address' not in data:
            return jsonify({'error': 'Address is required'}), 400
        
        # Extract parameters with defaults
        address = data['address']
        prepared_by = data.get('prepared_by', 'Brayden Fisher')
        prepared_by_company = data.get('prepared_by_company', 'Colliers International')
        prepared_by_address = data.get('prepared_by_address', '123 North 123 West Orem, UT 12345')
        prepared_for = data.get('prepared_for', 'Austin Shouse')
        prepared_for_company = data.get('prepared_for_company', 'UCCU Bank')
        prepared_for_address = data.get('prepared_for_address', '789 yellow street Provo, UT 12345')
        property_name = data.get('property_name', 'Property Report')
        property_type = data.get('property_type', 'Office')
        lot_area = data.get('lot_area', '')
        acres = data.get('acres', '')
        recorded_sale_date = data.get('recorded_sale_date', '')
        zoning = data.get('zoning', '')
        apn = data.get('apn', '')
        current_owner = data.get('current_owner', '')
        
        logger.info(f"Generating property report for: {address}")
        
        # Generate the report
        document_path = property_generator.process_single_property(
            address=address,
            prepared_by=prepared_by,
            prepared_by_company=prepared_by_company,
            prepared_by_address=prepared_by_address,
            prepared_for=prepared_for,
            prepared_for_company=prepared_for_company,
            prepared_for_address=prepared_for_address,
            property_name=property_name,
            property_type=property_type,
            lot_area=lot_area,
            acres=acres,
            recorded_sale_date=recorded_sale_date,
            zoning=zoning,
            apn=apn,
            current_owner=current_owner
        )
        
        # Return the download URL and success message
        filename = os.path.basename(document_path)
        download_url = f"/download_property_report/{filename}"
        
        return jsonify({
            'success': True,
            'message': 'Property report generated successfully',
            'document_path': document_path,
            'filename': filename,
            'download_url': download_url
        })
        
    except Exception as e:
        logger.error(f"Error generating property report: {e}")
        logger.error(traceback.format_exc())
        return jsonify({
            'error': f'Failed to generate property report: {str(e)}'
        }), 500

@app.route('/download_property_report/<filename>', methods=['GET'])
def download_property_report(filename):
    """Download a generated property report"""
    try:
        file_path = Path("property_reports") / filename
        if not file_path.exists():
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        logger.error(f"Error downloading file: {e}")
        return jsonify({'error': f'Failed to download file: {str(e)}'}), 500

@app.route('/process_comps', methods=['POST'])
def process_comps():
    """
    Process comparable properties from PDF and generate report
    
    Expected multipart form data:
    - pdf_file: PDF file containing comp data (required)
    - template_file: (optional) Word template file
    - property_data: (optional) JSON string with property information for combined report
    
    Example:
    curl -X POST http://localhost:5000/process_comps \
      -F "pdf_file=@comp.pdf" \
      -F "template_file=@comptemplate.docx" \
      -F "property_data={\"address\":\"123 Main St\"}"
    """
    if not comp_extractor:
        return jsonify({'error': 'Comp extractor not initialized'}), 500
    
    try:
        # Check if files are uploaded
        if not request.files:
            return jsonify({'error': 'PDF file is required. Please upload a PDF file.'}), 400
        
        if 'pdf_file' not in request.files:
            return jsonify({'error': 'PDF file is required'}), 400
        
        pdf_file = request.files['pdf_file']
        template_file = request.files.get('template_file')
        
        # Validate PDF file
        if pdf_file.filename == '':
            return jsonify({'error': 'No PDF file selected'}), 400
        
        if not pdf_file.filename.lower().endswith('.pdf'):
            return jsonify({'error': 'File must be a PDF'}), 400
        
        # Save uploaded files temporarily
        temp_dir = Path(tempfile.mkdtemp())
        pdf_path = temp_dir / "comp.pdf"
        template_path = temp_dir / "comptemplate.docx" if template_file else COMP_TEMPLATE_PATH
        
        pdf_file.save(pdf_path)
        if template_file:
            template_file.save(template_path)
        
        # Extract comps from PDF
        logger.info(f"Extracting comps from PDF: {pdf_path}")
        comps = comp_extractor.extract_comps_from_pdf(str(pdf_path))
        
        # Check for optional property data for combined report
        property_data = None
        if 'property_data' in request.form:
            try:
                property_data = json.loads(request.form['property_data'])
                logger.info("Property data received for combined report")
            except json.JSONDecodeError as e:
                logger.warning(f"Invalid property_data JSON: {e}")
                property_data = None
        
        if not comps:
            # Try to get more information about the PDF content
            try:
                import pdfplumber
                with pdfplumber.open(str(pdf_path)) as pdf:
                    full_text = ""
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if page_text:
                            full_text += page_text + "\n"
                    
                    # Log first 500 characters to help debug
                    preview = full_text[:500].replace('\n', '\\n')
                    logger.warning(f"PDF content preview: {preview}")
                    
                    # Check if the expected pattern exists
                    import re
                    comp_pattern = r'(\d+)\.\s+Sold Land (?:Property|Space)'
                    matches = re.findall(comp_pattern, full_text)
                    logger.warning(f"Found {len(matches)} potential comp entries with pattern")
                    
                    if not matches:
                        return jsonify({
                            'error': 'No comparable properties found in PDF',
                            'details': 'PDF does not contain expected "Sold Land Property" or "Sold Land Space" entries',
                            'content_preview': preview
                        }), 400
                    else:
                        return jsonify({
                            'error': 'No comparable properties found in PDF',
                            'details': f'Found {len(matches)} potential entries but parsing failed',
                            'content_preview': preview
                        }), 400
                        
            except Exception as debug_error:
                logger.error(f"Error during debug analysis: {debug_error}")
                return jsonify({'error': 'No comparable properties found in PDF'}), 400
        
        results = {}
        
        # Generate comp report
        logger.info(f"Generating comp report with {len(comps)} comps")
        comp_output_path = comp_extractor.replace_keywords_in_document(
            str(template_path), 
            comps
        )
        
        comp_filename = os.path.basename(comp_output_path)
        comp_download_url = f"/download_comp_report/{comp_filename}"
        
        results['comp_report'] = {
            'success': True,
            'document_path': comp_output_path,
            'filename': comp_filename,
            'download_url': comp_download_url,
            'comps_count': len(comps),
            'comps_data': [
                {
                    'comp_number': comp.comp_number,
                    'property_name': comp.property_name,
                    'address': comp.address,
                    'sale_price': comp.sale_price,
                    'sale_price_sf': comp.sale_price_sf,
                    'market': comp.market,
                    'sub_market': comp.sub_market
                }
                for comp in comps[:6]  # Return first 6 comps
            ]
        }
        
        # Generate property report if property data is provided
        if property_data and property_generator:
            try:
                logger.info("Generating property report with provided data")
                property_document_path = property_generator.process_single_property(**property_data)
                property_filename = os.path.basename(property_document_path)
                property_download_url = f"/download_property_report/{property_filename}"
                
                results['property_report'] = {
                    'success': True,
                    'document_path': property_document_path,
                    'filename': property_filename,
                    'download_url': property_download_url
                }
            except Exception as e:
                logger.error(f"Error generating property report: {e}")
                results['property_report'] = {
                    'success': False,
                    'error': str(e)
                }
        
        # Clean up temporary files
        if temp_dir and temp_dir.exists():
            shutil.rmtree(temp_dir)
        
        # Return success response
        if len(results) == 1:
            # Only comp report
            return jsonify({
                'success': True,
                'message': f'Processed {len(comps)} comparable properties',
                **results['comp_report']
            })
        else:
            # Combined report
            return jsonify({
                'success': True,
                'message': 'Combined report generated successfully',
                'results': results
            })
        
    except Exception as e:
        logger.error(f"Error processing comps: {e}")
        logger.error(traceback.format_exc())
        return jsonify({
            'error': f'Failed to process comps: {str(e)}'
        }), 500

@app.route('/download_comp_report/<filename>', methods=['GET'])
def download_comp_report(filename):
    """Download a generated comp report"""
    try:
        file_path = Path("property_reports") / filename
        if not file_path.exists():
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        logger.error(f"Error downloading file: {e}")
        return jsonify({'error': f'Failed to download file: {str(e)}'}), 500

@app.route('/generate_combined_report', methods=['POST'])
def generate_combined_report():
    """
    Generate both property report and comp report in one call
    
    Expected JSON payload:
    {
        "address": "501 N 730 W American Fork Ut 84003",
        "property_data": { ... },  // All property report parameters
        "pdf_path": "comp.pdf",    // Path to comp PDF
        "template_path": "comptemplate.docx"  // Optional
    }
    """
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'Request data is required'}), 400
        
        results = {}
        
        # Generate property report if address is provided
        if 'address' in data:
            property_data = data.get('property_data', {})
            property_data['address'] = data['address']
            
            if property_generator:
                document_path = property_generator.process_single_property(**property_data)
                filename = os.path.basename(document_path)
                download_url = f"/download_property_report/{filename}"
                results['property_report'] = {
                    'success': True,
                    'document_path': document_path,
                    'filename': filename,
                    'download_url': download_url
                }
            else:
                results['property_report'] = {
                    'success': False,
                    'error': 'Property generator not initialized'
                }
        
        # Process comps if PDF path is provided
        if 'pdf_path' in data:
            if comp_extractor:
                pdf_path = data['pdf_path']
                temp_dir = None
                
                # Check if pdf_path is a URL
                if pdf_path.startswith(('http://', 'https://')):
                    temp_dir = Path(tempfile.mkdtemp())
                    pdf_path = download_pdf_from_url(pdf_path, temp_dir)
                
                comps = comp_extractor.extract_comps_from_pdf(str(pdf_path))
                template_path = data.get('template_path', COMP_TEMPLATE_PATH)
                output_path = comp_extractor.replace_keywords_in_document(template_path, comps)
                
                # Clean up temporary files
                if temp_dir and temp_dir.exists():
                    shutil.rmtree(temp_dir)
                filename = os.path.basename(output_path)
                download_url = f"/download_comp_report/{filename}"
                
                results['comp_report'] = {
                    'success': True,
                    'document_path': output_path,
                    'filename': filename,
                    'download_url': download_url,
                    'comps_count': len(comps)
                }
            else:
                results['comp_report'] = {
                    'success': False,
                    'error': 'Comp extractor not initialized'
                }
        
        return jsonify({
            'success': True,
            'message': 'Combined report generation completed',
            'results': results
        })
        
    except Exception as e:
        logger.error(f"Error generating combined report: {e}")
        logger.error(traceback.format_exc())
        return jsonify({
            'error': f'Failed to generate combined report: {str(e)}'
        }), 500

@app.route('/list_reports', methods=['GET'])
def list_reports():
    """List all generated reports"""
    try:
        reports_dir = Path("property_reports")
        if not reports_dir.exists():
            return jsonify({'reports': []})
        
        reports = []
        for file_path in reports_dir.glob("*.docx"):
            stats = file_path.stat()
            reports.append({
                'filename': file_path.name,
                'size_bytes': stats.st_size,
                'created': datetime.fromtimestamp(stats.st_ctime).isoformat(),
                'modified': datetime.fromtimestamp(stats.st_mtime).isoformat()
            })
        
        # Sort by modification time (newest first)
        reports.sort(key=lambda x: x['modified'], reverse=True)
        
        return jsonify({
            'reports': reports,
            'total_count': len(reports)
        })
        
    except Exception as e:
        logger.error(f"Error listing reports: {e}")
        return jsonify({'error': f'Failed to list reports: {str(e)}'}), 500

@app.route('/delete_report/<filename>', methods=['DELETE'])
def delete_report(filename):
    """Delete a specific report"""
    try:
        file_path = Path("property_reports") / filename
        if not file_path.exists():
            return jsonify({'error': 'File not found'}), 404
        
        file_path.unlink()
        
        return jsonify({
            'success': True,
            'message': f'Report {filename} deleted successfully'
        })
        
    except Exception as e:
        logger.error(f"Error deleting file: {e}")
        return jsonify({'error': f'Failed to delete file: {str(e)}'}), 500

if __name__ == '__main__':
    # Initialize generators on startup
    initialize_generators()
    
    # Run the Flask app
    app.run(
        host='0.0.0.0',  # Allow external connections
        port=5000,
        debug=True
    ) 