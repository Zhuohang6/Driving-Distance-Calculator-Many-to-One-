import pandas as pd
import openrouteservice
import time

def geocode_address(client, address):
    """
    Use OpenRouteService's Pelias geocoding service to get (latitude, longitude).
    Returns (None, None) if not found.
    """
    try:
        response = client.pelias_search(text=address, size=1)
        
        if response and 'features' in response and len(response['features']) > 0:
            # Extract the first feature geometry
            coords = response['features'][0]['geometry']['coordinates']
            # coords is [lon, lat]
            lon = coords[0]
            lat = coords[1]
            return (lat, lon)
    except Exception as e:
        print(f"Geocoding error for '{address}': {e}")
    return (None, None)


def get_driving_distance_in_km(client, origin_lat, origin_lon, dest_lat, dest_lon):
    """
    Use OpenRouteService Directions API to get driving distance in kilometers 
    between (origin_lat, origin_lon) and (dest_lat, dest_lon).
    Returns None if something fails.
    """
    try:
        coords = [
            (origin_lon, origin_lat),
            (dest_lon, dest_lat)
        ]
        
        # Request driving directions
        # profile='driving-car' => car route
        # returns a route with distance in meters
        route = client.directions(coords, profile='driving-car',radiuses=[2000, 2000])
        
        # Check the structure for distance in the first route
        if route and 'routes' in route and len(route['routes']) > 0:
            distance_meters = route['routes'][0]['summary']['distance']
            distance_km = distance_meters / 1000.0
            return distance_km
    except Exception as e:
        print(f"Directions error: {e}")
    
    return None


def main():
    # 1. Your input file (Excel) with addresses
    input_excel = "addresses.xlsx"
    # 2. Your output Excel file
    output_excel = "output_with_distances.xlsx"
    
    # 3. Address as the single destination
    Fixed_address = 'Fixed address'
    
    # 4. Instantiate the OpenRouteService client with your API key
    #    Sign up at https://openrouteservice.org/ to get a key
    ors_api_key = "key"  # Replace with your actual API key
    client = openrouteservice.Client(key=ors_api_key)
    
    # 5. Geocode the destination address once (to reduce API calls)
    Fixed_lat, Fixed_lon = geocode_address(client, Fixed_address)
    if Fixed_lat is None or Fixed_lon is None:
        print("Error: Could not geocode the destination address.")
        return
    
    # 6. Read your input Excel
    df = pd.read_excel(input_excel)
    
    # Make sure your Excel has these columns: 
    # "Address", "City", "Region", "PostalCode", "CountryId"
    # Adjust accordingly if your column names differ.
    
    # 7. Prepare a new column to store the distance
    df["Distance (km)"] = None
    
    # 8. Loop through each row, build the full address string, 
    #    geocode, get distance, store result
    for index, row in df.iterrows():
        # Build address string
        address_parts = [
            str(row["Address"]) if not pd.isnull(row["Address"]) else "",
            str(row["City"]) if not pd.isnull(row["City"]) else "",
            str(row["Region"]) if not pd.isnull(row["Region"]) else "",
            str(row["PostalCode"]) if not pd.isnull(row["PostalCode"]) else "",
            str(row["CountryId"]) if not pd.isnull(row["CountryId"]) else ""
        ]
        origin_address = ", ".join([part.strip() for part in address_parts if part.strip()])
        
        # Geocode origin address
        origin_lat, origin_lon = geocode_address(client, origin_address)
        
        # If we got valid coords, get driving distance
        distance_km = None
        if origin_lat is not None and origin_lon is not None:
            distance_km = get_driving_distance_in_km(
                client,
                origin_lat, origin_lon,
                Fixed_lat, Fixed_lon
            )
        
        # Store distance (could be None if not found)
        df.at[index, "Distance (km)"] = distance_km
        
        # (Optional) To avoid hitting rate limits for large data, you can add a slight delay:
        # time.sleep(0.2)  # 200ms delay per request, for instance
    
    # 9. Write out results to a new Excel file
    df.to_excel(output_excel, index=False)
    print(f"Distances written to '{output_excel}'.")


if __name__ == "__main__":
    main()
