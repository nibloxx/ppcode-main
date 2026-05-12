#!/usr/bin/env python3
"""
Test script to verify API server functionality
"""

import requests
import json
import time

def test_health():
    """Test health endpoint"""
    print("🔍 Testing health endpoint...")
    try:
        response = requests.get("http://localhost:5000/health")
        if response.status_code == 200:
            data = response.json()
            print(f"✅ Health check passed")
            print(f"   Status: {data['status']}")
            print(f"   Property Generator: {data['property_generator']}")
            print(f"   Comp Extractor: {data['comp_extractor']}")
            return True
        else:
            print(f"❌ Health check failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Health check error: {e}")
        return False

def test_reinitialize():
    """Test reinitialize endpoint"""
    print("\n🔄 Testing reinitialize endpoint...")
    try:
        response = requests.post("http://localhost:5000/reinitialize")
        if response.status_code == 200:
            data = response.json()
            print(f"✅ Reinitialize successful")
            print(f"   Property Generator: {data['property_generator']}")
            print(f"   Comp Extractor: {data['comp_extractor']}")
            return True
        else:
            print(f"❌ Reinitialize failed: {response.status_code}")
            print(f"   Response: {response.text}")
            return False
    except Exception as e:
        print(f"❌ Reinitialize error: {e}")
        return False

def test_property_report():
    """Test property report generation"""
    print("\n🏠 Testing property report generation...")
    try:
        data = {
            "address": "501 N 730 W American Fork Ut 84003",
            "prepared_by": "Test User",
            "prepared_by_company": "Test Company",
            "property_type": "Office"
        }
        
        response = requests.post(
            "http://localhost:5000/generate_property_report",
            json=data,
            headers={"Content-Type": "application/json"}
        )
        
        if response.status_code == 200:
            result = response.json()
            print(f"✅ Property report generated successfully")
            print(f"   Filename: {result['filename']}")
            return result['filename']
        else:
            print(f"❌ Property report generation failed: {response.status_code}")
            print(f"   Response: {response.text}")
            return None
    except Exception as e:
        print(f"❌ Property report error: {e}")
        return None

def test_process_comps():
    """Test comp processing"""
    print("\n📊 Testing comp processing...")
    try:
        data = {
            "pdf_path": "comp.pdf",
            "template_path": "comptemplate.docx"
        }
        
        response = requests.post(
            "http://localhost:5000/process_comps",
            json=data,
            headers={"Content-Type": "application/json"}
        )
        
        if response.status_code == 200:
            result = response.json()
            print(f"✅ Comp processing successful")
            print(f"   Comps found: {result['comps_count']}")
            print(f"   Filename: {result['filename']}")
            return result['filename']
        else:
            print(f"❌ Comp processing failed: {response.status_code}")
            print(f"   Response: {response.text}")
            return None
    except Exception as e:
        print(f"❌ Comp processing error: {e}")
        return None

def test_list_reports():
    """Test list reports endpoint"""
    print("\n📋 Testing list reports...")
    try:
        response = requests.get("http://localhost:5000/list_reports")
        if response.status_code == 200:
            data = response.json()
            print(f"✅ List reports successful")
            print(f"   Total reports: {data['total_count']}")
            for report in data['reports'][:3]:  # Show first 3
                print(f"   - {report['filename']}")
            return True
        else:
            print(f"❌ List reports failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ List reports error: {e}")
        return False

def main():
    print("🧪 Property Report API Test Suite")
    print("=" * 50)
    
    # Wait a moment for server to be ready
    print("⏳ Waiting for server to be ready...")
    time.sleep(2)
    
    # Test health
    if not test_health():
        print("\n❌ Server is not responding. Make sure it's running on http://localhost:5000")
        return
    
    # Test reinitialize
    test_reinitialize()
    
    # Test property report
    property_filename = test_property_report()
    
    # Test comp processing
    comp_filename = test_process_comps()
    
    # Test list reports
    test_list_reports()
    
    print("\n" + "=" * 50)
    print("🎉 Test suite completed!")
    
    if property_filename:
        print(f"📄 Property report: {property_filename}")
    if comp_filename:
        print(f"📊 Comp report: {comp_filename}")
    
    print("\n📝 Next steps:")
    print("   1. Use Postman to test the endpoints")
    print("   2. Check the generated files in property_reports/")
    print("   3. Download files using the download endpoints")

if __name__ == "__main__":
    main() 