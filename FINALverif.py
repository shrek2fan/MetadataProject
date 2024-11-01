import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
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
    }
    return approved_subjects




# Load the city dataset into a dictionary structure with safe handling for non-string values
def load_city_data(city_dataset_path):
    city_data = pd.read_excel(city_dataset_path)

    def safe_strip(value):
        # Only strip if value is a string; otherwise, return an empty string or the value as-is
        return str(value).strip() if isinstance(value, str) else ''

    city_info = {}
    for _, row in city_data.iterrows():
        city_info[(safe_strip(row['ES_City']).lower(), 'spanish')] = {
            'country': safe_strip(row['ES_Country']),
            'state': safe_strip(row['ES_State']),
            'coordinates': safe_strip(row["CITIES' LAT_LONG COORDINATES"])
        }
        city_info[(safe_strip(row['EN_City']).lower(), 'english')] = {
            'country': safe_strip(row['EN_Country']),
            'state': safe_strip(row['EN_State']),
            'coordinates': safe_strip(row["CITIES' LAT_LONG COORDINATES"])
        }
    return city_info

# Define validation functions for each column
def is_valid_digital_identifier(value):
    if not isinstance(value, str):
        return False, "Value is not a string"
    if not value.startswith("Ms0004") or not value.endswith(".pdf"):
        return False, "Identifier does not follow expected format"
    return True, "Valid"


def is_valid_box_folder(value):
    """
    Validates the box folder format.

    Returns:
        Tuple[bool, str]: (True, "Valid") if the format is correct,
                          (False, <Error message>) otherwise.
    """
    if not isinstance(value, str):
        return False, "Box Folder value is not a string"
    if not re.match(r'^\d{2}_\d{2}$', value):
        return False, "Box Folder format is incorrect, expected 'XX_XX' with two digits before and after the underscore"
    return True, "Valid"


def is_valid_collection_name(value, language="English"):
    """
    Validates the collection name based on the language.

    Args:
        value (str): The collection name to validate.
        language (str): The language context ("English" or "Spanish").

    Returns:
        Tuple[bool, str]: (True, "Valid") if the name is correct,
                          (False, <Error message>) otherwise.
    """
    if not isinstance(value, str):
        return False, "Collection Name is not a string"
    
    expected_name = "Correspondencia de la familia Amador, 1856-1949" if language.lower() == "spanish" else "Amador Family Correspondence, 1856-1949"
    if value != expected_name:
        return False, f"Collection Name does not match expected value for {language}. Expected '{expected_name}' but got '{value}'"
    
    return True, "Valid"



def is_valid_date(value):
    """
    Validates that the date is in either 'YYYY-MM-DD' or 'YYYY-MM' format.

    Returns:
        Tuple[bool, str]: (True, "Valid") if the date format is correct,
                          (False, <Error message>) otherwise.
    """
    if pd.isna(value):
        return True, "Valid (empty)"
    
    # Check for 'YYYY-MM-DD' format
    try:
        pd.to_datetime(value, format='%Y-%m-%d', errors='raise')
        return True, "Valid"
    except (ValueError, TypeError):
        pass
    
    # Check for 'YYYY-MM' format
    try:
        pd.to_datetime(value, format='%Y-%m', errors='raise')
        return True, "Valid"
    except (ValueError, TypeError):
        return False, "Date format is invalid. Expected 'YYYY-MM-DD' or 'YYYY-MM'"


def is_valid_year(value):
    """
    Validates the year.

    Returns:
        Tuple[bool, str]: (True, "Valid") if the year is correct,
                          (False, <Error message>) otherwise.
    """
    if not isinstance(value, int):
        return False, "Year is not an integer"
    if not (1000 <= value <= 9999):
        return False, "Year is out of valid range (1000-9999)"
    return True, "Valid"



def is_valid_subject_lcsh(value, language="english"):
    if not isinstance(value, str):
        return False, "Invalid type: Expected a string"

    # Split terms based on separator and strip any extra whitespace from each term
    terms = [term.strip() for term in value.split("[|]")]
    
    # Check for any whitespace issues between terms and separators
    if "[|]" in value and " " in value:
        return False, "Whitespace found around separator or terms"

    # Verify each term exists in the vocabulary set for the specified language
    invalid_terms = [term for term in terms if term not in approved_subjects[language]]
    if invalid_terms:
        return False, f"Terms not found in vocabulary: {invalid_terms}"

    # If all terms are valid, return True
    return True, "Valid subject terms"


def check_name_format(value):
    if not isinstance(value, str):
        return "missing"
    stripped_name = value.strip()
    return "valid" if stripped_name in authorized_names else "missing"

def is_valid_city_related(row, city_column, country_column, state_column, coord_column, language):
    # Get city, country, state, and coordinates, handling NaNs
    city = str(row.get(city_column, '')).strip().lower() if pd.notna(row.get(city_column)) else ''
    country = row.get(country_column, '').strip() if pd.notna(row.get(country_column)) else ''
    state = row.get(state_column, '').strip() if pd.notna(row.get(state_column)) else ''
    coordinates = row.get(coord_column, '').strip() if pd.notna(row.get(coord_column)) else ''
    
    # Skip validation if city is empty or "no data"
    if not city or city == "no data":
        return True, "City data missing or marked as 'no data'"

    # Retrieve expected data for the city
    city_key = (city, language)
    expected_data = city_info.get(city_key)
    
    # City not found in dataset
    if not expected_data:
        return False, f"City '{city}' not found in dataset for language '{language}'"

    # Country and State Validation
    valid_country = country == expected_data['country'] if country else True
    valid_state = state == expected_data['state'] if state else True

    # Detailed logging for country/state mismatches
    if not valid_country:
        return False, f"Country mismatch: Expected '{expected_data['country']}', found '{country}'"
    if not valid_state:
        return False, f"State mismatch: Expected '{expected_data['state']}', found '{state}'"

    # Coordinate Validation
    expected_coords = expected_data['coordinates']
    coord_sets = coordinates.split("[|]")

    # Ensure 1 or 2 coordinate sets exist and match format
    if len(coord_sets) > 2:
        return False, "Coordinate format error: More than two coordinate sets found. Both GEOLOC_SCITY columns should be highlighted."

    # Check if coordinates match exactly as expected, in the correct format
    for i, actual_coords in enumerate(coord_sets):
        if actual_coords.strip() != expected_coords:
            return False, f"Coordinate set {i + 1} does not match expected value '{expected_coords}'. Highlight GEOLOC_SCITY columns."

    # If the cell contains two sets of coordinates, ensure the correct separator is present
    if len(coord_sets) == 2 and coordinates != f"{coord_sets[0].strip()}[|]{coord_sets[1].strip()}":
        return False, "Coordinate format error: Missing or incorrect '[|]' separator for dual coordinates. Highlight GEOLOC_SCITY columns."

    return True, "Location data matches expected values"


def load_authorized_names(names_dataset_path):
    names_data = pd.read_excel(names_dataset_path, usecols=[0])
    return set(names_data['PEOPLE'].dropna().str.strip())

# Column validation rules
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

location_validation_rules = {
    "SENDERS_CITY": lambda row: is_valid_city_related(row, 'SENDERS_CITY', 'SENDERS_COUNTRY', 'SENDERS_STATE', 'GEOLOC_SCITY', 'english'),
    "ES..SENDERS_CITY": lambda row: is_valid_city_related(row, 'ES..SENDERS_CITY', 'ES..SENDERS_COUNTRY', 'ES..SENDERS_STATE', 'ES..GEOLOC_SCITY', 'spanish'),
    "ADDRESSEES_CITY": lambda row: is_valid_city_related(row, 'ADDRESSEES_CITY', 'ADDRESSEES_COUNTRY', 'ADDRESSEES_STATE', 'GEOLOC_SCITY', 'english'),
    "ES..ADDRESSEES_CITY": lambda row: is_valid_city_related(row, 'ES..ADDRESSEES_CITY', 'ES..ADDRESSEES_COUNTRY', 'ES..ADDRESSEES_STATE', 'ES..GEOLOC_SCITY', 'spanish')
}

# Load data for validation
approved_subjects = load_approved_subjects("SUBJECT_LCSH.xlsx")
city_info = load_city_data("Maybeee.xlsx")
authorized_names = load_authorized_names("CVPeople.xlsx")


def verify_file(input_file, output_file):
    wb = load_workbook(input_file)
    ws = wb["OA_Descriptive metadata"]
    df = pd.read_excel(input_file, sheet_name="OA_Descriptive metadata")
    
    for idx, row in df.iterrows():
        # Apply general validation rules
        for col_name, validation_func in column_validation_rules.items():
            if col_name in df.columns:
                value = row[col_name]
                try:
                    is_valid, message = validation_func(value)  # Unpacking the result
                    if not is_valid:
                        col_idx = df.columns.get_loc(col_name) + 1
                        ws.cell(row=idx + 2, column=col_idx).fill = highlight_fill_red
                        print(f"Failed validation: {col_name} at row {idx + 2} - Reason: {message}")
                    else:
                        print(f"Validation successful: {col_name} at row {idx + 2}")
                except Exception as e:
                    print(f"Error validating {col_name} at row {idx + 2}: {e}")

        # Apply location-specific validation
        for col_name, location_func in location_validation_rules.items():
            if col_name in df.columns:
                try:
                    is_valid, message = location_func(row)  # Unpacking the result
                    if not is_valid:
                        col_idx = df.columns.get_loc(col_name) + 1
                        ws.cell(row=idx + 2, column=col_idx).fill = highlight_fill_red
                        print(f"Failed location validation: {col_name} at row {idx + 2} - Reason: {message}")
                    else:
                        print(f"Location validation successful: {col_name} at row {idx + 2}")
                except Exception as e:
                    print(f"Error in location validation for {col_name} at row {idx + 2}: {e}")

    wb.save(output_file)
    print(f"Verification completed. Output saved as {output_file}")



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
