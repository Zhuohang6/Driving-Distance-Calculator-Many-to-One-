# Driving-Distance-Calculator-Many-to-One-
This script calculates driving distances (in kilometers) between a fixed address and multiple addresses listed in an Excel file. It uses the OpenRouteService API for geocoding and routing.

## Features
- Geocodes addresses to latitude/longitude.
- Computes driving distances using OpenRouteService.
- Outputs results to a new Excel file.

## Prerequisites
- Python 3.7+
- OpenRouteService API key (sign up at [OpenRouteService](https://openrouteservice.org/)).

## Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/distance-calculator.git
   cd distance-calculator

2. Install the required Python libraries:
pip install -r requirements.txt

3. Add your API key in calculate_distance.py.

4. Edit your fixed_address.

5. Prepare your input file (addresses.xlsx) with the following columns:
Address
City
Region
PostalCode
Country

6. Run the script:
python calculate_distance.py


# Input Requirements for FullAddressEntry.py:
-Open Excel
-Create a new column with the exact name "FullAddress".
-Enter full addresses in each row following the format:
-Street Address, City, State, ZIP Code, Country
-Save the file as .xlsx








