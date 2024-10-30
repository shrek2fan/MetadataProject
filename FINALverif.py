import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl import Workbook
import re
import argparse
import os

# Define fill styles for highlighting mistakes
highlight_fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
highlight_fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# Load the approved SUBJECT_LCSH vocabulary with separate English and Spanish terms
def load_approved_subjects(vocabulary_file):
    approved_df = pd.read_excel(vocabulary_file)
    english_terms = approved_df.iloc[:, 0].str.strip().tolist()
    spanish_terms = approved_df.iloc[:, 1].str.strip().tolist()
    approved_subjects = {
        "english": set(english_terms),
        "spanish": set(spanish_terms),
        "english_to_spanish": dict(zip(english_terms, spanish_terms)),
        "spanish_to_english": dict(zip(spanish_terms, english_terms))
    }
    return approved_subjects

# Load the city dataset into a dictionary structure
def load_city_data(city_dataset_path):
    city_data = pd.read_excel(city_dataset_path)
    
    def safe_strip(value):
        return str(value).strip() if pd.notna(value) else ''
    
    city_info = {}
    for _, row in city_data.iterrows():
        city_info[(safe_strip(row['ES_City']), 'spanish')] = {
            'country': safe_strip(row['ES_Country']),
            'state': safe_strip(row['ES_State']),
            'coordinates': safe_strip(row["CITIES' LAT_LONG COORDINATES"])
        }
        city_info[(safe_strip(row['EN_City']), 'english')] = {
            'country': safe_strip(row['EN_Country']),
            'state': safe_strip(row['EN_State']),
            'coordinates': safe_strip(row["CITIES' LAT_LONG COORDINATES"])
        }
    return city_info

# Load the authorized names dataset to enforce exact matches in FROM and TO fields
def load_authorized_names(names_dataset_path):
    names_data = pd.read_excel(names_dataset_path, usecols=[0])
    return set(names_data['PEOPLE'].dropna().str.strip())

# Load data for validation
approved_subjects = load_approved_subjects("SUBJECT_LCSH.xlsx")
city_info = load_city_data("Maybeee.xlsx")
authorized_names = load_authorized_names("CVPeople.xlsx")

# Define validation functions for each column
def is_valid_digital_identifier(value):
    return isinstance(value, str) and value.startswith("Ms0004") and value.endswith(".pdf")

def is_valid_box_folder(value):
    return isinstance(value, str) and bool(re.match(r'\d{2}_\d{2}', value))

def is_valid_collection_name(value, language="English"):
    return value == ("Correspondencia de la familia Amador, 1856-1949" if language == "Spanish" else "Amador Family Correspondence, 1856-1949")

def is_valid_date(value):
    if pd.isna(value):
        return True
    try:
        pd.to_datetime(value, format='%Y-%m-%d', errors='raise')
        return True
    except ValueError:
        try:
            pd.to_datetime(value, format='%Y-%m', errors='raise')
            return True
        except ValueError:
            return False

def is_valid_year(value):
    return isinstance(value, int) and 1000 <= value <= 9999

def is_valid_subject_lcsh(value, language="english"):
    if not isinstance(value, str):
        return False
    terms = [term.strip() for term in value.split("[|]")]
    return all(term in approved_subjects[language] for term in terms)

# Updated function to handle two cases for authorized names
def check_name_format(value):
    """
    Returns:
    - "missing" if the name is not found in the authorized names list.
    - "format_error" if the name is found but has formatting issues.
    - "valid" if the name is correctly formatted.
    """
    if not isinstance(value, str):
        return "missing"
    
    stripped_name = value.strip()
    if stripped_name not in authorized_names:
        return "missing"
    
    if stripped_name != value:
        # Formatting issues (e.g., extra spaces)
        return "format_error"
    
    return "valid"

def is_valid_city_related(row, city_column, country_column, state_column, coord_column, language):
    # Safely get city, country, state, and coordinates values, handling NaNs and ensuring they are strings
    city = row.get(city_column, '')
    city = str(city).strip().lower() if isinstance(city, str) and city.lower() != "no data" else ''
    
    # Log the exact city value read from the row for debugging
    print(f"Checking city in row: '{city}'")

    # Skip validation if city is empty after stripping
    if not city:
        print(f"Skipping row due to missing or 'no data' city in column '{city_column}'.")
        return True  

    # Retrieve expected data from the city_info dictionary
    city_key = (city, language)
    expected_data = city_info.get(city_key)
    
    # If the city is not in the dataset, raise an error
    if not expected_data:
        print(f"City '{city}' not found in the dataset for language '{language}'.")
        return False  # City not in dataset

    # Retrieve and sanitize comparison values for the current row, skipping "no data" fields
    country = str(row.get(country_column, '')).strip() if isinstance(row.get(country_column), str) and row.get(country_column).lower() != "no data" else ''
    state = str(row.get(state_column, '')).strip() if isinstance(row.get(state_column), str) and row.get(state_column).lower() != "no data" else ''
    coordinates = str(row.get(coord_column, '')).strip() if isinstance(row.get(coord_column), str) and row.get(coord_column).lower() != "no data" else ''

    # Parse expected and actual coordinates to floats, handling errors
    try:
        expected_coords = [float(x) for x in expected_data['coordinates'].split(",")]
        actual_coords = [float(x) for x in coordinates.split(",")] if coordinates else []
        valid_coords = len(expected_coords) == len(actual_coords) and all(
            round(ec, 6) == round(ac, 6) for ec, ac in zip(expected_coords, actual_coords)
        ) if coordinates else True  # If coordinates are blank, skip comparison
    except ValueError:
        print(f"Invalid coordinate format in row with city '{city}'.")
        valid_coords = False  # Set to False if parsing fails

    # Check if each of the fields matches the expected data
    valid_country = (country == expected_data['country']) if country else True  # Skip if country is blank
    valid_state = (state == expected_data['state']) if state else True  # Skip if state is blank
    
    # Debugging output for each comparison
    print(f"Expected values for city '{city}': Country='{expected_data['country']}', State='{expected_data['state']}', Coordinates='{expected_data['coordinates']}'")
    print(f"Actual values: Country='{country}', State='{state}', Coordinates='{coordinates}'")
    print(f"  Match results -> Country: {valid_country}, State: {valid_state}, Coordinates: {valid_coords}")

    # If any of the fields with values do not match, return False to flag this row
    if not (valid_country and valid_state and valid_coords):
        print(f"Mismatch detected for city '{city}':")
        print(f"  Expected: {{Country: '{expected_data['country']}', State: '{expected_data['state']}', Coordinates: '{expected_data['coordinates']}'}}")
        print(f"  Found: {{Country: '{country}', State: '{state}', Coordinates: '{coordinates}'}}")
        return False  # Highlight mismatch in the calling function

    # Return True if all non-"no data" values match
    return True



# Column validation rules to map columns to their respective validation functions
column_validation_rules = {
    "DIGITAL_IDENTIFIER": is_valid_digital_identifier,
    "ES..DIGITAL_IDENTIFIER": is_valid_digital_identifier,
    "BOX_FOLDER": is_valid_box_folder,
    "ES..BOX_FOLDER": is_valid_box_folder,
    "COLLECTION_NAME": lambda x: is_valid_collection_name(x, language="English"),
    "ES..COLLECTION_NAME": lambda x: is_valid_collection_name(x, language="Spanish"),
    "DATE": is_valid_date,
    "ES..DATE": is_valid_date,
    "YEAR": is_valid_year,
    "ES..YEAR": is_valid_year,
    "SUBJECT_LCSH": lambda x: is_valid_subject_lcsh(x, language="english"),
    "ES..SUBJECT_LCSH": lambda x: is_valid_subject_lcsh(x, language="spanish"),
    "FROM": check_name_format,
    "ES..FROM": check_name_format,
    "TO": check_name_format,
    "ES..TO": check_name_format
}

# Main verification function with detailed logging
def verify_file(input_file, output_file):
    # Load the input Excel file
    df = pd.read_excel(input_file, sheet_name="OA_Descriptive metadata")

    # Print column names for verification
    print("Columns in the DataFrame:", df.columns)

    # Define the columns for city data to check
    city_columns = ['SENDERS_CITY', 'ES..SENDERS_CITY']

    # Initialize workbook and worksheet to save results
    wb = Workbook()
    ws = wb.active
    ws.title = "Verification Results"
    
    # Highlight fill color for mismatched cells
    highlight_fill_red = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    # Iterate over each row for verification
    for idx, row in df.iterrows():
        # Find the city from either of the specified columns
        city = None
        for city_col in city_columns:
            if city_col in df.columns:
                city = str(row.get(city_col, '')).strip().lower()
                if city:
                    break

        # Skip if no city found in the relevant columns
        if not city:
            print(f"Skipping row {idx + 2} due to missing city in columns {city_columns}.")
            continue

        # Assuming language is 'english' for this example, adjust as needed
        language = 'english'

        # Fetch expected data for the city from the city_info dataset
        expected_data = city_info.get((city, language))
        if expected_data:
            expected_coords = expected_data['coordinates']
            
            # Verify coordinates in columns GEOLOC_SCITY and ES..GEOLOC_SCITY
            geoloc_columns = ['GEOLOC_SCITY', 'ES..GEOLOC_SCITY']
            for geoloc_col in geoloc_columns:
                if geoloc_col in df.columns:
                    actual_coords = str(row.get(geoloc_col, '')).strip()
                    if actual_coords != expected_coords:
                        cell = f"{geoloc_col}{idx + 2}"
                        ws[cell].fill = highlight_fill_red
                        print(f"Mismatch in {geoloc_col} at row {idx + 2}: Expected '{expected_coords}', Found '{actual_coords}'")
        else:
            print(f"City '{city}' not found in the dataset for language '{language}'.")

    # Save the output file
    wb.save(output_file)
    print(f"Verification completed. Mistakes highlighted in {output_file}")


# Argument parser to take file input and generate verified output
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Verify an Excel file against validation rules.")
    parser.add_argument("file_name", help="The name of the Excel file to verify")
    args = parser.parse_args()
    
    # Generate output file name with "Verified_" prefix
    input_file = args.file_name
    output_file = f"Verified_{os.path.basename(input_file)}"
    
    # Run verification
    verify_file(input_file, output_file)