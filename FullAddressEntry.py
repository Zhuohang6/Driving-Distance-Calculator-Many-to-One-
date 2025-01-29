import pandas as pd
import googlemaps
import os
import time
import random
from geopy.distance import geodesic
 
# Load Google API Key from environment variable
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
 
if not GOOGLE_API_KEY:
    raise ValueError("Google API Key is missing! Set GOOGLE_API_KEY as an environment variable.")
 
# Base coordinates
BASE_ADDRESS_COORDINATES = (1, -1)
 
# Initialize Google Maps client
gmaps = googlemaps.Client(key=GOOGLE_API_KEY)
 
def clean_string(value):
    """Converts values to string, handling NaN values."""
    return str(value).strip() if pd.notna(value) else ''
 
def extract_address_components(full_address):
    """Splits FullAddress into meaningful components safely."""
    parts = [p.strip() for p in full_address.split(",") if p.strip()]  # Remove empty values and extra spaces
 
    # Ensure we have at least 3 components (Address, City, Country)
    if len(parts) < 3:
        print(f"Warning: Incomplete address found: '{full_address}'")
        return {"Address": "", "City": "", "Region": "", "PostalCode": "", "CountryId": ""}
 
    return {
        "Address": parts[0] if len(parts) > 0 else '',
        "City": parts[1] if len(parts) > 1 else '',
        "Region": parts[2] if len(parts) > 2 else '',
        "PostalCode": parts[3] if len(parts) > 3 else '',
        "CountryId": parts[4] if len(parts) > 4 else ''
    }
 
def retry_with_backoff(retries=5, base_delay=1):
    """Retries API calls using exponential backoff."""
    for i in range(retries):
        delay = base_delay * (2 ** i) + random.uniform(0, 0.1)
        print(f"Retrying in {delay:.2f} seconds...")
        time.sleep(delay)
 
def geocode_address(address, retries=3):
    """Geocode an address using Google Maps API with retries."""
    if not address.strip():
        print("Skipping empty address.")
        return None
   
    for attempt in range(retries):
        try:
            print(f"Geocoding: {address}")
            geocode_result = gmaps.geocode(address)
            if geocode_result:
                location = geocode_result[0]['geometry']['location']
                return location['lat'], location['lng']
        except googlemaps.exceptions.ApiError as e:
            print(f"Google API Error: {e}")
        except Exception as e:
            print(f"Unexpected geocoding error: {e}")
        retry_with_backoff()
    return None
 
def calculate_distance(origin, destination):
    """Calculate driving distance between two points."""
    try:
        directions_result = gmaps.directions(origin, destination, mode="driving")
        if directions_result:
            return directions_result[0]['legs'][0]['distance']['value'] / 1000  # Convert to km
    except Exception as e:
        print(f"Google Maps distance API failed: {e}")
    return geodesic(origin, destination).kilometers  # Fallback to geodesic distance
 
def validate_coordinates(coords, max_distance=5000):
    """Validate coordinates by checking distance from base."""
    if not coords:
        return False
    distance_from_base = geodesic(BASE_ADDRESS_COORDINATES, coords).kilometers
    return (-90 <= coords[0] <= 90 and -180 <= coords[1] <= 180 and distance_from_base <= max_distance)
 
def process_addresses(input_file, output_file, geocode_max_distance=5000):
    """Process addresses and save distance calculations."""
    if not os.path.isfile(input_file):
        print(f"Error: File '{input_file}' not found.")
        return
 
    df = pd.read_excel(input_file)
    results = []
 
    for _, row in df.iterrows():
        full_address = clean_string(row.get('FullAddress'))
        components = extract_address_components(full_address)
        structured_address = f"{components['Address']}, {components['City']}, {components['Region']}, {components['PostalCode']}, {components['CountryId']}"
       
        destination_coords = geocode_address(structured_address)
       
        if destination_coords and validate_coordinates(destination_coords, geocode_max_distance):
            distance = calculate_distance(BASE_ADDRESS_COORDINATES, destination_coords)
        else:
            distance = "NA"
       
        results.append({
            'FullAddress': full_address,
            'Extracted Address': components['Address'],
            'PostalCode': components['PostalCode'],
            'CountryId': components['CountryId'],
            'Distance (km)': distance
        })
 
        time.sleep(random.uniform(1, 2))  # Randomized delay to avoid rate limits
 
    # Save results with timestamp
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    output_filename = f"{output_file.rstrip('.xlsx')}_{timestamp}.xlsx"
    pd.DataFrame(results).to_excel(output_filename, index=False)
    print(f"Results saved to '{output_filename}'.")
 
if __name__ == "__main__":
    input_file = "addresses.xlsx"
    output_file = "distances.xlsx"
    process_addresses(input_file, output_file)
