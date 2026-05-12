# 🚀 Quick Start Guide - Property Report API

This guide will help you get the API server running with your existing `both4.py` and `comp2.py` files.

## 📋 Prerequisites

Make sure you have these files in your directory:
- ✅ `both4.py` - Property report generator
- ✅ `comp2.py` - Comparable properties extractor
- ✅ `template.docx` - Property report template
- ✅ `comptemplate.docx` - Comp report template
- ✅ `comp.pdf` - Sample comp PDF (optional, for testing)

## 🔧 Installation

### Option 1: Automatic Installation (Recommended)
```bash
python install_dependencies.py
```

### Option 2: Manual Installation
```bash
pip install -r requirements_minimal.txt
```

## 🚀 Start the Server

### Option 1: Easy Start
```bash
python start_api.py
```

### Option 2: Direct Start
```bash
python api_server.py
```

The server will start on `http://localhost:5000`

## 🧪 Test the API

Run the test suite to verify everything is working:
```bash
python test_api.py
```

## 📱 Use with Postman

1. **Import the collection:**
   - Open Postman
   - Click "Import"
   - Select `Property_Report_API.postman_collection.json`

2. **Test the endpoints:**
   - Start with `Health Check` to verify server is running
   - Use `Reinitialize Generators` if needed
   - Try `Generate Property Report` with your data
   - Try `Process Comps` with your PDF

## 🔗 API Endpoints

### Health & Status
- `GET /health` - Check server status
- `POST /reinitialize` - Reinitialize generators

### Property Reports (from both4.py)
- `POST /generate_property_report` - Generate property reports
- `GET /download_property_report/<filename>` - Download reports

### Comparable Properties (from comp2.py)
- `POST /process_comps` - Process comps from PDF
- `GET /download_comp_report/<filename>` - Download comp reports

### Utility
- `GET /list_reports` - List all generated reports
- `DELETE /delete_report/<filename>` - Delete reports

## 📝 Example Usage

### Generate Property Report
```bash
curl -X POST http://localhost:5000/generate_property_report \
  -H "Content-Type: application/json" \
  -d '{
    "address": "501 N 730 W American Fork Ut 84003",
    "prepared_by": "Brayden Fisher",
    "prepared_by_company": "Colliers International",
    "property_type": "Office"
  }'
```

### Process Comps
```bash
curl -X POST http://localhost:5000/process_comps \
  -H "Content-Type: application/json" \
  -d '{
    "pdf_path": "comp.pdf",
    "template_path": "comptemplate.docx"
  }'
```

## 🔍 Troubleshooting

### "Comp extractor not initialized" Error
1. Check the health endpoint: `GET /health`
2. If comp_extractor is false, call: `POST /reinitialize`
3. Check the server logs for initialization errors

### Missing Files Error
Make sure these files exist:
- `both4.py`
- `comp2.py`
- `template.docx`
- `comptemplate.docx`

### Import Errors
If you get import errors:
1. Install dependencies: `pip install -r requirements_minimal.txt`
2. Check Python version (3.8+ required)
3. Try the installation script: `python install_dependencies.py`

## 📁 File Structure
```
ppcode/
├── both4.py                                    # ✅ Your property generator
├── comp2.py                                    # ✅ Your comp extractor
├── template.docx                               # ✅ Property template
├── comptemplate.docx                          # ✅ Comp template
├── comp.pdf                                   # ✅ Sample comp PDF
├── api_server.py                              # 🆕 Flask API server
├── start_api.py                               # 🆕 Easy startup script
├── test_api.py                                # 🆕 Test suite
├── install_dependencies.py                    # 🆕 Dependency installer
├── requirements_minimal.txt                   # 🆕 Dependencies
└── Property_Report_API.postman_collection.json # 🆕 Postman collection
```

## 🎯 Next Steps

1. **Test with Postman** - Import the collection and try all endpoints
2. **Customize templates** - Modify `template.docx` and `comptemplate.docx`
3. **Add your data** - Replace sample data with your actual property information
4. **Integrate with your app** - Use the API endpoints in your applications

## 📞 Support

If you encounter issues:
1. Check the server logs for error messages
2. Run `python test_api.py` to identify problems
3. Verify all required files are present
4. Ensure dependencies are installed correctly

The API server now uses your existing `both4.py` and `comp2.py` files directly, so all your existing functionality is preserved! 