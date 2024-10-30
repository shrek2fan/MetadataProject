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




# Load the city dataset into a dictionary structure
def load_city_data(city_dataset_path):
    city_data = pd.read_excel(city_dataset_path)
    city_info = {}
    for _, row in city_data.iterrows():
        city_info[(row['ES_City'].strip().lower(), 'spanish')] = {
            'country': row['ES_Country'].strip(),
            'state': row['ES_State'].strip(),
            'coordinates': row["CITIES' LAT_LONG COORDINATES"].strip()
        }
        city_info[(row['EN_City'].strip().lower(), 'english')] = {
            'country': row['EN_Country'].strip(),
            'state': row['EN_State'].strip(),
            'coordinates': row["CITIES' LAT_LONG COORDINATES"].strip()
        }
    return city_info

# Define validation functions for each column
def is_valid_digital_identifier(value):
    return isinstance(value, str) and value.startswith("Ms0004") and value.endswith(".pdf")

def is_valid_box_folder(value):
    return isinstance(value, str) and bool(re.match(r'\d{2}_\d{2}', value))

def is_valid_collection_name(value, language="English"):
    return value == ("Correspondencia de la familia Amador, 1856-1949" if language == "Spanish" else "Amador Family Correspondence, 1856-1949")

def is_valid_date(value):
    try:
        pd.to_datetime(value, format='%Y-%m-%d', errors='raise')
        return True
    except (ValueError, TypeError):
        return False

def is_valid_year(value):
    return isinstance(value, int) and 1000 <= value <= 9999

def is_valid_subject_lcsh(value, language="english"):
    if not isinstance(value, str):
        return False
    terms = [term.strip() for term in value.split("[|]")]
    return all(term in approved_subjects[language] for term in terms)

def check_name_format(value):
    if not isinstance(value, str):
        return "missing"
    stripped_name = value.strip()
    return "valid" if stripped_name in authorized_names else "missing"

def is_valid_city_related(row, city_column, country_column, state_column, coord_column, language):
    city = str(row.get(city_column, '')).strip().lower() if pd.notna(row.get(city_column)) else ''
    if not city or city == "no data":
        return True  # Skips validation if city is missing or set to 'no data'
    
    # Retrieve expected data for the city
    city_key = (city, language)
    expected_data = city_info.get(city_key)
    
    if not expected_data:
        return False  # City not in dataset
    
    # Coordinate and location checks
    valid_country = row[country_column].strip() == expected_data['country'] if pd.notna(row.get(country_column)) else True
    valid_state = row[state_column].strip() == expected_data['state'] if pd.notna(row.get(state_column)) else True
    valid_coords = row[coord_column].strip() == expected_data['coordinates'] if pd.notna(row.get(coord_column)) else True
    
    return valid_country and valid_state and valid_coords


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
                    if not validation_func(value):
                        col_idx = df.columns.get_loc(col_name) + 1
                        ws.cell(row=idx + 2, column=col_idx).fill = highlight_fill_red
                        print(f"Failed validation: {col_name} at row {idx + 2}")
                except Exception as e:
                    print(f"Error in {col_name} at row {idx + 2}: {e}")

        # Apply location-specific validation
        for col_name, location_func in location_validation_rules.items():
            if col_name in df.columns:
                try:
                    if not location_func(row):
                        col_idx = df.columns.get_loc(col_name) + 1
                        ws.cell(row=idx + 2, column=col_idx).fill = highlight_fill_red
                        print(f"Failed location validation: {col_name} at row {idx + 2}")
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
