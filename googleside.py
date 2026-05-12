import requests
from datetime import datetime

def get_property_images(address):
    """
    Get 3 images of the property: 2 Street View + 1 Aerial
    
    Args:
        address: The address to get images for
    """
    api_key = 'AIzaSyCl6Oc03tJ-MkQEXMc84pF9lXURvPLPmHU'
    
    # Convert address to coordinates
    geocoding_params = {
        "address": address,
        "key": api_key
    }
    
    try:
        # Get coordinates from address
        geocoding_response = requests.get(
            "https://maps.googleapis.com/maps/api/geocode/json", 
            params=geocoding_params,
            timeout=10
        )
        geocoding_response.raise_for_status()
        geocoding_data = geocoding_response.json()
        
        if geocoding_data.get("status") != "OK":
            print(f"Error: Could not find address - {geocoding_data.get('status')}")
            return
        
        if not geocoding_data.get("results"):
            print(f"Error: No results found for address: {address}")
            return
        
        location = geocoding_data["results"][0]["geometry"]["location"]
        lat, lng = location["lat"], location["lng"]
        print(f"Coordinates found: {lat}, {lng}")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 1. Get Aerial/Satellite Image
        print("Fetching aerial image...")
        aerial_params = {
            "center": f"{lat},{lng}",
            "zoom": "18",
            "size": "1920x1920",
            "maptype": "satellite",  # Aerial view
            "key": api_key,
            "markers": f"{lat},{lng}"
        }
        
        aerial_response = requests.get(
            "https://maps.googleapis.com/maps/api/staticmap", 
            params=aerial_params,
            timeout=30
        )
        aerial_response.raise_for_status()
        
        aerial_filename = f"aerial_{timestamp}.jpg"
        with open(aerial_filename, "wb") as f:
            f.write(aerial_response.content)
        print(f"✓ Aerial image saved as: {aerial_filename}")
        
        # 2. Get Street View Image - Looking directly at the property
        print("Fetching Street View image...")
        
        # Street View API parameters - positioned to look at the address
        street_view_params = {
            "size": "600x500",
            "location": address,     # Use the address directly
            "pitch": "0",           # Level view
            "fov": "90",            # Field of view
            "key": api_key
        }
        
        # Get Street View image
        street_view_response = requests.get(
            "https://maps.googleapis.com/maps/api/streetview", 
            params=street_view_params,
            timeout=30
        )
        street_view_response.raise_for_status()
        
        # Check if Street View is available
        if street_view_response.headers.get('content-type', '').startswith('image/'):
            street_view_filename = f"street_view_{timestamp}.jpg"
            with open(street_view_filename, "wb") as f:
                f.write(street_view_response.content)
            print(f"✓ Street view saved as: {street_view_filename}")
        else:
            print(f"✗ Street View not available for this address")
        
        print(f"\n🎉 Property images completed!")
        print(f"📁 Files saved:")
        print(f"   • {aerial_filename} (aerial view)")
        print(f"   • street_view_{timestamp}.jpg (facing the property)")
        
    except requests.RequestException as e:
        print(f"Error: Failed to fetch images - {e}")
    except Exception as e:
        print(f"Error: {e}")

def main():
    print("Property Image Fetcher")
    print("=====================")
    print("This will get 2 images:")
    print("• 1 Aerial/satellite view (bird's eye)")
    print("• 1 Street View photo (facing the house)")
    print()
    
    address = input("Enter property address: ").strip()
    
    if not address:
        print("Error: Please enter an address")
        return
    
    get_property_images(address)

if __name__ == "__main__":
    main()


#OG 