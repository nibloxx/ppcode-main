#!/usr/bin/env python3
"""
Simple startup script for the Property Report API Server
"""

import sys
import os
from pathlib import Path

def check_dependencies():
    """Check if required files exist"""
    required_files = [
        "both4.py",
        "comp2.py", 
        "template.docx",
        "comptemplate.docx"
    ]
    
    missing_files = []
    for file in required_files:
        if not Path(file).exists():
            missing_files.append(file)
    
    if missing_files:
        print("❌ Missing required files:")
        for file in missing_files:
            print(f"   - {file}")
        print("\nPlease ensure all required files are in the current directory.")
        return False
    
    print("✅ All required files found")
    return True

def check_directories():
    """Create required directories if they don't exist"""
    directories = ["property_reports", "property_reports/images", "property_reports/comp_images"]
    
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)
    
    print("✅ Directories created/verified")

def main():
    print("🚀 Property Report API Server")
    print("=" * 40)
    
    # Check dependencies
    if not check_dependencies():
        sys.exit(1)
    
    # Create directories
    check_directories()
    
    print("\n📋 Starting API server...")
    print("   - Server will run on: http://localhost:5000")
    print("   - Health check: http://localhost:5000/health")
    print("   - Press Ctrl+C to stop the server")
    print("\n" + "=" * 40)
    
    try:
        # Import and run the API server
        from api_server import app
        
        app.run(
            host='0.0.0.0',
            port=5000,
            debug=True
        )
        
    except KeyboardInterrupt:
        print("\n\n🛑 Server stopped by user")
    except ImportError as e:
        print(f"\n❌ Import error: {e}")
        print("Please install dependencies: pip install -r requirements.txt")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error starting server: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 