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
    approved_df = pd.read_excel(vocabulary_file, usecols=[0, 1])  # Load only columns A and B
    approved_subjects = {
        "spanish": set(approved_df.iloc[:, 0].dropna().str.strip()),  # Spanish terms from column A
        "english": set(approved_df.iloc[:, 1].dropna().str.strip()),  # English terms from column B
    }
    print("Debug: Approved subjects loaded:", approved_subjects)  # Log loaded terms for confirmation
    return approved_subjects


# Load the city dataset into a dictionary structure with safe handling for non-string values
def load_city_data(city_dataset_path):
    city_data = pd.read_excel(city_dataset_path)

    def safe_strip(value):
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


def is_valid_city_related(row, city_column, country_column, state_column, coord_column, language):
    """
    Validates the city, country, state, and coordinates data, ensuring the coordinates match
    the expected values for the given city in the specified language.

    Parameters:
    - row (pd.Series): The row of data being validated.
    - city_column (str): Column name for the city.
    - country_column (str): Column name for the country.
    - state_column (str): Column name for the state.
    - coord_column (str): Column name for the coordinates.
    - language (str): Language ('english' or 'spanish') to select the appropriate city data.

    Returns:
    - (bool, str, str): Tuple indicating if validation passed, the column to highlight if not,
      and an error message.
    """
    # Retrieve and clean city, country, state, and coordinates
    city = str(row.get(city_column, '')).strip().lower() if pd.notna(row.get(city_column)) else ''
    country = row.get(country_column, '').strip() if pd.notna(row.get(country_column)) else ''
    state = row.get(state_column, '').strip() if pd.notna(row.get(state_column)) else ''
    coordinates = row.get(coord_column, '').strip() if pd.notna(row.get(coord_column)) else ''

    # Skip validation if city is empty or marked as "no data"
    if not city or city == "no data":
        return True, "", "City data missing or marked as 'no data'"

    # Fetch expected data for the city from city_info dictionary
    city_key = (city, language)
    expected_data = city_info.get(city_key)

    # If city is not in the dataset, highlight in yellow
    if not expected_data:
        return False, "yellow", f"City '{city}' not found in dataset for language '{language}'"

    # Validate country and state if provided
    if country and country != expected_data['country']:
        return False, "red", f"Country mismatch: Expected '{expected_data['country']}', found '{country}'"
    
    if state and state != expected_data['state']:
        return False, "red", f"State mismatch: Expected '{expected_data['state']}', found '{state}'"

    # Validate coordinates
    expected_coords = expected_data['coordinates']
    coord_sets = coordinates.split("[|]")

    # Ensure coordinate format is correct
    if len(coord_sets) > 2:
        return False, "red", "Coordinate format error: More than two coordinate sets found."

    # Check that at least one of the sets matches expected coordinates for each city
    matched_coords = [actual_coords.strip() for actual_coords in coord_sets if actual_coords.strip() == expected_coords]

    if not matched_coords:
        return False, "red", f"No matching coordinates found for '{city}' with expected value '{expected_coords}'."

    # If dual coordinates are provided, ensure correct separator format
    if len(coord_sets) == 2 and coordinates != f"{coord_sets[0].strip()}[|]{coord_sets[1].strip()}":
        return False, "red", "Coordinate format error: Incorrect '[|]' separator for dual coordinates."

    return True, "", "Location data matches expected values"


# Adjust location validation rules to include the language parameter
location_validation_rules = {
    "SENDERS_CITY": lambda row: is_valid_city_related(row, 'SENDERS_CITY', 'SENDERS_COUNTRY', 'SENDERS_STATE', 'GEOLOC_SCITY', 'english'),
    "ES..SENDERS_CITY": lambda row: is_valid_city_related(row, 'ES..SENDERS_CITY', 'ES..SENDERS_COUNTRY', 'ES..SENDERS_STATE', 'ES..GEOLOC_SCITY', 'spanish'),
    "ADDRESSEES_CITY": lambda row: is_valid_city_related(row, 'ADDRESSEES_CITY', 'ADDRESSEES_COUNTRY', 'ADDRESSEES_STATE', 'GEOLOC_SCITY', 'english'),
    "ES..ADDRESSEES_CITY": lambda row: is_valid_city_related(row, 'ES..ADDRESSEES_CITY', 'ES..ADDRESSEES_COUNTRY', 'ES..ADDRESSEES_STATE', 'ES..GEOLOC_SCITY', 'spanish')
}




def validate_digital_identifier(value, previous_identifier=None):
    """
    Validate the DIGITAL_IDENTIFIER format for either 'Ms0004_XX_XX_XX.pdf' or 'Ms0071_XX_XX_XX.pdf' pattern.
    The format must follow 'Ms0004_<BoxNumber>_<FolderNumber>_<LetterNumber>.pdf' or 'Ms0071_<BoxNumber>_<FolderNumber>_<LetterNumber>.pdf'.
    The BoxNumber and FolderNumber should remain consistent in sequence, and
    LetterNumber should increment by 1 from the previous row's LetterNumber.

    Parameters:
    - value (str): The identifier to validate.
    - previous_identifier (tuple): The box, folder, and letter numbers of the previous identifier.

    Returns:
    - (bool, str, str): Validation status, fill color ('red' or 'yellow'), and message.
    """
    if not isinstance(value, str):
        return False, "red", "Invalid type: Expected a string"

    # Check format: Ms0004_XX_XX_XX.pdf or Ms0071_XX_XX_XX.pdf
    match = re.match(r"^(Ms0004|Ms0071)_(\d{2})_(\d{2})_(\d{2})\.pdf$", value)
    if not match:
        return False, "red", "Incorrect format. Expected 'Ms0004_XX_XX_XX.pdf' or 'Ms0071_XX_XX_XX.pdf' where XX are two-digit numbers"

    # Extract BoxNumber, FolderNumber, and LetterNumber
    _, box_number, folder_number, letter_number = match.groups()

    # Check if this is the first identifier in the series
    if previous_identifier is None:
        # Initialize tracking with the first identifier's box and folder number, starting with LetterNumber = 01
        if letter_number != "01":
            return False, "red", "First letter number must start with 01"
        return True, (box_number, folder_number, int(letter_number)), "Valid"

    # Unpack previous identifier
    prev_box, prev_folder, prev_letter = previous_identifier

    # Verify that BoxNumber and FolderNumber match the previous identifier's values
    if box_number != prev_box or folder_number != prev_folder:
        return False, "red", f"Box or folder number mismatch. Expected '{prev_box}_{prev_folder}'"

    # Ensure LetterNumber increments by 1
    if int(letter_number) != prev_letter + 1:
        return False, "red", f"Letter number must increment sequentially. Expected {str(prev_letter + 1).zfill(2)}"

    # If all checks pass, update previous identifier tracking and mark as valid
    return True, (box_number, folder_number, int(letter_number)), "Valid"


def is_valid_box_folder(value):
    if not isinstance(value, str):
        return False, None, "Box Folder value is not a string"
    if not re.match(r'^\d{2}_\d{2}$', value):
        return False, None, "Box Folder format is incorrect, expected 'XX_XX' with two digits before and after the underscore"
    return True, None, "Valid"

def is_valid_collection_name(value, language="English"):
    if not isinstance(value, str):
        return False, None, "Collection Name is not a string"
    
    expected_name = "Correspondencia de la familia Amador, 1856-1949" if language.lower() == "spanish" else "Amador Family Correspondence, 1856-1949"
    if value != expected_name:
        return False, None, f"Collection Name does not match expected value for {language}. Expected '{expected_name}' but got '{value}'"
    
    return True, None, "Valid"

def is_valid_date(value):
    if pd.isna(value):
        return True, None, "Valid (empty)"
    try:
        pd.to_datetime(value, format='%Y-%m-%d', errors='raise')
        return True, None, "Valid"
    except (ValueError, TypeError):
        pass
    try:
        pd.to_datetime(value, format='%Y-%m', errors='raise')
        return True, None, "Valid"
    except (ValueError, TypeError):
        return False, None, "Date format is invalid. Expected 'YYYY-MM-DD' or 'YYYY-MM'"

def is_valid_year(value):
    if not isinstance(value, int):
        return False, None, "Year is not an integer"
    if not (1000 <= value <= 9999):
        return False, None, "Year is out of valid range (1000-9999)"
    return True, None, "Valid"

def validate_name_field(value, authorized_names):
    """
    Validates if a name in the FROM or TO fields matches an entry in the authorized names dataset.
    If the name is missing from the dataset entirely, it highlights the cell in yellow.
    If the name exists in the dataset but is not in the correct format, it highlights the cell in red.

    Parameters:
    - value (str): The name to validate.
    - authorized_names (set): Set of authorized names loaded from the dataset.

    Returns:
    - (bool, str, str): Validation status, fill color ('red', 'yellow', or None), and message.
    """
    known_unknown_values = {
        "Unknown sender", "Remitente desconocido",
        "Unknown recipient", "Destinatario desconocido"
    }

    # Clean up the input value
    cleaned_value = value.strip()
    
    print(f"Debug: Starting search for '{cleaned_value}' in authorized names...")

    # Check if the value is in the known "unknown" set
    if cleaned_value in known_unknown_values:
        print(f"Debug: '{cleaned_value}' recognized as a known unknown value. Validation passed.")
        return True, None, "Valid (unknown value)"
    
    # Check if cleaned_value matches an authorized name exactly
    if cleaned_value in authorized_names:
        print(f"Debug: '{cleaned_value}' found in authorized names with correct format. Validation passed.")
        return True, None, "Valid"
    
    # Check if a similar name exists but does not match exactly
    matching_names = [name for name in authorized_names if name.lower() == cleaned_value.lower()]
    if matching_names:
        print(f"Debug: '{cleaned_value}' found in authorized names but format is incorrect. Expected '{matching_names[0]}'.")
        return False, "red", f"Format error: Expected '{matching_names[0]}'"
    
    # Name does not exist in the authorized names dataset at all
    print(f"Debug: '{cleaned_value}' not found in authorized names dataset.")
    return False, "yellow", "Name not found in dataset"


def is_valid_subject_lcsh(value, approved_subjects, language="english"):
    if not isinstance(value, str):
        return False, None, "Invalid type: Expected a string"

    terms = [term.strip() for term in value.split("[|]")]
    print(f"Debug: Starting validation for SUBJECT_LCSH terms '{terms}' in language '{language}'...")

    invalid_terms = []
    for term in terms:
        if term in approved_subjects[language]:
            print(f"Debug: Term '{term}' found in vocabulary for '{language}'. Validation passed.")
        else:
            print(f"Debug: Term '{term}' NOT found in vocabulary for '{language}'. Validation failed.")
            invalid_terms.append(term)

    if invalid_terms:
        return False, "yellow", f"Terms not found in vocabulary: {invalid_terms}"
    return True, None, "Valid"



def load_authorized_names(names_dataset_path):
    names_data = pd.read_excel(names_dataset_path, usecols=[0], header=None)
    authorized_names = set(names_data[0].dropna().str.strip())
    print("Debug: Authorized names loaded from dataset:", authorized_names)
    return authorized_names

# Column validation rules
column_validation_rules = {
    "DIGITAL_IDENTIFIER": validate_digital_identifier,
    "ES..DIGITAL_IDENTIFIER": validate_digital_identifier,
    "BOX_FOLDER": is_valid_box_folder,
    "ES..BOX_FOLDER": is_valid_box_folder,
    "COLLECTION_NAME": lambda x: is_valid_collection_name(x, language="English"),
    "ES..COLLECTION_NAME": lambda x: is_valid_collection_name(x, language="Spanish"),
    "DATE": is_valid_date,
    "ES..DATE": is_valid_date,
    "YEAR": is_valid_year,
    "ES..YEAR": is_valid_year,
    "SUBJECT_LCSH": lambda x: is_valid_subject_lcsh(x, approved_subjects, language="english"),
    "ES..SUBJECT_LCSH" : lambda x: is_valid_subject_lcsh(x, approved_subjects, language="spanish"),
    "FROM": lambda x: validate_name_field(x, authorized_names),
    "ES..FROM": lambda x: validate_name_field(x, authorized_names),
    "TO": lambda x: validate_name_field(x, authorized_names),
    "ES..TO": lambda x: validate_name_field(x, authorized_names),
}

location_validation_rules = {
    "SENDERS_CITY": lambda row: is_valid_city_related(row, 'SENDERS_CITY', 'SENDERS_COUNTRY', 'SENDERS_STATE', 'GEOLOC_SCITY', 'english'),
    "ES..SENDERS_CITY": lambda row: is_valid_city_related(row, 'ES..SENDERS_CITY', 'ES..SENDERS_COUNTRY', 'ES..SENDERS_STATE', 'ES..GEOLOC_SCITY', 'spanish'),
    "ADDRESSEES_CITY": lambda row: is_valid_city_related(row, 'ADDRESSEES_CITY', 'ADDRESSEES_COUNTRY', 'ADDRESSEES_STATE', 'GEOLOC_SCITY', 'english'),
    "ES..ADDRESSEES_CITY": lambda row: is_valid_city_related(row, 'ES..ADDRESSEES_CITY', 'ES..ADDRESSEES_COUNTRY', 'ES..ADDRESSEES_STATE', 'ES..GEOLOC_SCITY', 'spanish')
}



approved_subjects = load_approved_subjects("SUBJECT_LCSH.xlsx")
city_info = load_city_data("Maybeee.xlsx")
authorized_names = load_authorized_names("CVPeople.xlsx")
print("Loaded authorized names:", authorized_names)

# The `verify_file` function and main script setup remain the same, using `column_validation_rules`.



def verify_file(input_file, output_file):
    wb = load_workbook(input_file)
    ws = wb["OA_Descriptive metadata"]
    df = pd.read_excel(input_file, sheet_name="OA_Descriptive metadata")

    previous_identifier = None  # Initialize for tracking previous DIGITAL_IDENTIFIER

    for idx, row in df.iterrows():
        # Validate non-location columns
        for col_name, validation_func in column_validation_rules.items():
            if col_name in df.columns:
                value = row[col_name]
                try:
                    # Special handling for DIGITAL_IDENTIFIER to track sequence
                    if col_name in ["DIGITAL_IDENTIFIER", "ES..DIGITAL_IDENTIFIER"]:
                        is_valid, result, message = validate_digital_identifier(value, previous_identifier)
                        if is_valid:
                            previous_identifier = result  # Update previous_identifier for next iteration
                            print(f"Validation successful: {col_name} at row {idx + 2}")
                        else:
                            col_idx = df.columns.get_loc(col_name) + 1
                            ws.cell(row=idx + 2, column=col_idx).fill = highlight_fill_red if result == "red" else highlight_fill_yellow
                            print(f"Failed validation: {col_name} at row {idx + 2} - Reason: {message}")
                    else:
                        is_valid, color, message = validation_func(value)
                        if is_valid:
                            print(f"Validation successful: {col_name} at row {idx + 2}")
                        else:
                            col_idx = df.columns.get_loc(col_name) + 1
                            ws.cell(row=idx + 2, column=col_idx).fill = highlight_fill_red if color == "red" else highlight_fill_yellow
                            print(f"Failed validation: {col_name} at row {idx + 2} - Reason: {message}")
                except Exception as e:
                    print(f"Error validating {col_name} at row {idx + 2}: {e}")

        # Validate location-related columns
        for loc_col_name, location_func in location_validation_rules.items():
            if loc_col_name in df.columns:
                try:
                    is_valid, highlight_col, message = location_func(row)
                    if is_valid:
                        print(f"Location validation successful: {loc_col_name} at row {idx + 2}")
                    else:
                        col_idx = df.columns.get_loc(highlight_col) + 1
                        ws.cell(row=idx + 2, column=col_idx).fill = highlight_fill_red
                        print(f"Failed location validation: {loc_col_name} at row {idx + 2} - Reason: {message}")
                except Exception as e:
                    print(f"Error in location validation for {loc_col_name} at row {idx + 2}: {e}")

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
