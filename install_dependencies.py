#!/usr/bin/env python3
"""
Dependency installation script with fallback options
"""

import subprocess
import sys
import os
from pathlib import Path

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description} completed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} failed:")
        print(f"   Error: {e.stderr}")
        return False

def install_package(package, description=None):
    """Install a single package with error handling"""
    if description is None:
        description = f"Installing {package}"
    
    return run_command(f"pip install {package}", description)

def main():
    print("🚀 Property Report API - Dependency Installer")
    print("=" * 50)
    
    # Check if pip is available
    if not run_command("pip --version", "Checking pip"):
        print("❌ pip is not available. Please install pip first.")
        sys.exit(1)
    
    # Upgrade pip first
    run_command("python -m pip install --upgrade pip", "Upgrading pip")
    
    # Install setuptools first (this often fixes googlemaps issues)
    print("\n📦 Installing core dependencies...")
    if not install_package("setuptools==68.2.2", "Installing setuptools"):
        print("⚠️  setuptools installation failed, trying without version...")
        install_package("setuptools", "Installing setuptools (latest)")
    
    # Install Flask and core web dependencies
    print("\n🌐 Installing web framework...")
    install_package("flask==2.3.3", "Installing Flask")
    install_package("flask-cors==4.0.0", "Installing Flask-CORS")
    
    # Install OpenAI
    print("\n🤖 Installing AI dependencies...")
    install_package("openai==1.3.0", "Installing OpenAI")
    
    # Install Google Maps with fallback
    print("\n🗺️  Installing mapping dependencies...")
    if not install_package("googlemaps==4.9.0", "Installing Google Maps"):
        print("⚠️  Google Maps installation failed, trying alternative...")
        if not install_package("googlemaps", "Installing Google Maps (latest)"):
            print("❌ Google Maps installation failed. You may need to install it manually.")
    
    # Install data processing
    print("\n📊 Installing data processing...")
    install_package("pandas==2.0.3", "Installing Pandas")
    install_package("requests==2.31.0", "Installing Requests")
    
    # Install document processing
    print("\n📄 Installing document processing...")
    install_package("python-docx==0.8.11", "Installing python-docx")
    
    # Install PDF processing
    print("\n📑 Installing PDF processing...")
    install_package("pdfplumber==0.9.0", "Installing pdfplumber")
    install_package("PyMuPDF==1.23.7", "Installing PyMuPDF")
    
    # Install geocoding
    print("\n🌍 Installing geocoding...")
    install_package("geopy==2.3.0", "Installing geopy")
    
    print("\n" + "=" * 50)
    print("✅ Dependency installation completed!")
    print("\n📋 Next steps:")
    print("   1. Make sure you have the required files:")
    print("      - both4.py")
    print("      - comp2.py")
    print("      - template.docx")
    print("      - comptemplate.docx")
    print("   2. Start the API server:")
    print("      python start_api.py")
    print("\n🔧 If you encounter any issues:")
    print("   - Try installing packages individually")
    print("   - Check your Python version (3.8+ recommended)")
    print("   - Ensure you have proper permissions")

if __name__ == "__main__":
    main() 