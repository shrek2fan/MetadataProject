import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import argparse
import os

# Define fill style for highlighting mistakes
highlight_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

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

def is_valid_city_related(row, city_column, country_column, state_column, coord_column, language):
    city = row[city_column].strip()
    if not city:
        return True  # Consider valid if no city is provided

    city_key = (city, language)
    expected_data = city_info.get(city_key)
    
    if not expected_data:
        return False  # City not in dataset

    valid_country = row[country_column].strip() == expected_data['country']
    valid_state = row[state_column].strip() == expected_data['state']
    valid_coords = row[coord_column].strip() == expected_data['coordinates']
    
    return valid_country and valid_state and valid_coords

def is_authorized_name(value):
    return value.strip() in authorized_names if isinstance(value, str) else True

# Map columns to their validation functions
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
    "FROM": is_authorized_name,
    "ES..FROM": is_authorized_name,
    "TO": is_authorized_name,
    "ES..TO": is_authorized_name
}

# Main verification function
def verify_file(transformed_file, output_file):
    xls = pd.ExcelFile(transformed_file)
    wb = load_workbook(transformed_file)
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        ws = wb[sheet_name]
        
        for column_name, validation_func in column_validation_rules.items():
            if column_name in df.columns:
                for idx, value in enumerate(df[column_name]):
                    if not validation_func(value):
                        cell = ws[f"{chr(65 + df.columns.get_loc(column_name))}{idx + 2}"]
                        cell.fill = highlight_fill
        
        # Apply city-related validation for each row
        for idx, row in df.iterrows():
            if not is_valid_city_related(row, 'City', 'Country', 'State', 'Coordinates', 'english'):
                for col in ['City', 'Country', 'State', 'Coordinates']:
                    cell = ws[f"{chr(65 + df.columns.get_loc(col))}{idx + 2}"]
                    cell.fill = highlight_fill
    
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
