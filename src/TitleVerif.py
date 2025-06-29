import pandas as pd
import numpy as np
import re
import os
import argparse
import unicodedata
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parents[1] / "data"

# Compiling regular expressions used throughout the script
box_folder_regex = re.compile(r'[^\d_]')  # Regex to clean up BOX_FOLDER, allows only digits and underscores
collection_number_regex = re.compile(r'[^\w\s]')  # Example regex for COLLECTION_NUMBER, allows alphanumeric characters
title_regex = re.compile(r'[^\w\s,.]')  # Allows letters, numbers, spaces, periods, and commas
clean_invitation_regex = re.compile(r'\binvitaci[oó]n\b', re.IGNORECASE)  # Regex to catch invitation typos in Spanish

# Argument parsing to accept the file name as input
parser = argparse.ArgumentParser(description="Process an Excel file")
parser.add_argument("file_name", help="The name of the Excel file with extension (e.g., Test5.xlsx)")
args = parser.parse_args()

# Get the file name and create the file path
file_name = args.file_name
file_path = os.path.join(os.getcwd(), file_name)

if not os.path.exists(file_path):
    raise FileNotFoundError(f"The file {file_name} does not exist in the current directory.")

# Load spreadsheet with multiple sheets
xls = pd.ExcelFile(file_path)

# Dictionary to hold cleaned DataFrames
cleaned_sheets = {}

# Load the known proper names from the Excel file
proper_names_df = pd.read_excel(DATA_DIR / 'ProperNames.xlsx')
proper_names_list = proper_names_df['PROPER AUTHORIZED NAMES'].str.strip().tolist()  # Create a list of proper names

def ensure_series(col):
    """Ensures the input is a Pandas Series."""
    if isinstance(col, pd.Series):
        return col
    elif isinstance(col, list) or isinstance(col, np.ndarray):
        return pd.Series(col)
    else:
        return pd.Series([col])


def clean_digital_identifier(df, col_name):
    """
    Cleans the DIGITAL_IDENTIFIER columns by ensuring the format is correct and
    numbers increment sequentially starting at row 2.

    Parameters:
    - df (pd.DataFrame): The DataFrame containing the column to be cleaned.
    - col_name (str): The name of the column to clean.

    Returns:
    - pd.Series: A cleaned column with proper DIGITAL_IDENTIFIER formatting.
    """
    col = df[col_name].fillna('').astype(str)
    
    # Determine the collection number from valid rows
    valid_row = col[col.str.contains(r'^Ms\d{4}_')].iloc[0]
    collection_number = valid_row.split('_')[0]
    
    # Ensure at least one valid row exists
    if collection_number not in ['Ms0004', 'Ms0071']:
        raise ValueError(f"Unexpected collection number in column '{col_name}': '{collection_number}'.")

    # Extract box and folder numbers
    try:
        _, box_number, folder_number, _ = valid_row.split('_')
    except ValueError:
        raise ValueError(f"Invalid format in first valid row: '{valid_row}'.")

    # Generate sequential letter numbers starting at 01
    def generate_identifier(index):
        letter_number = str(index + 1).zfill(2)  # Sequential number with leading zeros
        return f"{collection_number}_{box_number}_{folder_number}_{letter_number}.pdf"

    # Create cleaned column
    cleaned_col = pd.Series([generate_identifier(i) for i in range(len(col))], index=col.index)
    
    return cleaned_col





def create_full_folder_file_path(digital_identifier_col):
    digital_identifier_col = pd.Series(digital_identifier_col).fillna('').astype(str)
    
    def format_path(identifier):
        match = re.match(r'Ms0004_(\d{2})_(\d{2})_(\d{2})\.pdf', identifier)
        if match:
            box_number, folder_number, letter_number = match.groups()
            return f"/Box_{box_number}/{box_number}_{folder_number}/Ms0004_{box_number}_{folder_number}_{letter_number}.pdf"
        else:
            return identifier  # Return the original if it doesn't match the expected format
    
    return digital_identifier_col.apply(format_path)

def normalize_accents(text):
    """Normalize Spanish accents for proper display of words like 'invitación'."""
    if isinstance(text, str):
        # Normalize and ensure correct encoding for Spanish words
        text = unicodedata.normalize('NFKC', text)
    return text

def clean_title(col, language="Spanish"):
    col = ensure_series(col).astype(str).fillna('')  # Ensure it's a Pandas Series

    # Define a list of proper names to avoid changing
    proper_names_list = ['Maria', 'Clotilde', 'Fausto', 'Manuel', 'Adela']  # Add more names here

    def clean_value(x):
        if pd.isnull(x):
            return x
        x = str(x).strip()

        # Define words that should not be capitalized (prepositions, articles, etc.)
        lowercase_words = ['a', 'de', 'para', 'por', 'en', 'con', 'y', 'o', 'una']

        def capitalize_names(text):
            words = text.split()
            capitalized_words = []
            for i, word in enumerate(words):
                # Skip capitalization for proper names that should not be changed
                if word in proper_names_list:
                    capitalized_words.append(word)
                # Capitalize first word, proper nouns, but not connectors/prepositions
                elif i == 0 or word.lower() not in lowercase_words:
                    capitalized_words.append(word.capitalize())
                else:
                    capitalized_words.append(word.lower())
            return " ".join(capitalized_words)

        # Apply capitalization
        x = capitalize_names(x)

        return x

    return col.apply(clean_value)



def clean_title_english(col):
    col = ensure_series(col).astype(str).fillna('')  # Ensure it's a Pandas Series

    # Define a list of proper names to avoid changing
    proper_names_list = ['Maria', 'Clotilde', 'Fausto', 'Manuel', 'Adela']  # Add more names here

    def clean_value(x):
        if pd.isnull(x):
            return x
        x = str(x).strip()

        # Define words that should not be capitalized (prepositions, articles, etc.)
        lowercase_words = ['a', 'of', 'for', 'by', 'in', 'on', 'with', 'and', 'or', 'the']

        def capitalize_names(text):
            words = text.split()
            capitalized_words = []
            for i, word in enumerate(words):
                # Skip capitalization for proper names that should not be changed
                if word in proper_names_list:
                    capitalized_words.append(word)
                # Capitalize first word, proper nouns, but not connectors/prepositions
                elif i == 0 or word.lower() not in lowercase_words:
                    capitalized_words.append(word.capitalize())
                else:
                    capitalized_words.append(word.lower())
            return " ".join(capitalized_words)

        # Apply capitalization rules without translation
        x = capitalize_names(x)

        return x

    return col.apply(clean_value)








def clean_series(series_col):
    series_col = ensure_series(series_col)  # Ensure it's a Pandas Series
    return series_col.ffill().str.strip()  # Fill down and strip whitespaces


def clean_box_folder(col):
    col = ensure_series(col)  # Ensure it's a Pandas Series
    # Replace any non-numeric characters and ensure the separator is an underscore
    return col.str.replace(r'[^0-9]', '_', regex=True).str.strip()


def clean_collection_name(col, language="English"):
    """Fills the column with the appropriate collection name in the specified language."""
    col = ensure_series(col)  # Ensure it's a Pandas Series

    if language == "Spanish":
        value = "Correspondencia de la familia Amador, 1856-1949"
    else:
        value = "Amador Family Correspondence, 1856-1949"

    # Fill the entire column with the same value
    return pd.Series([value] * len(col), index=col.index)


def clean_collection_number(col):
    """Fills the entire column with the constant 'Ms0004'."""
    col = ensure_series(col)  # Make sure it's a Pandas Series
    return pd.Series(["Ms0004"] * len(col), index=col.index)


#Helper function
def ensure_series(col):
    """Ensure the input is a Pandas Series."""
    if isinstance(col, pd.Series):
        return col
    elif isinstance(col, list) or isinstance(col, np.ndarray):
        return pd.Series(col)
    else:
        return pd.Series([col])



# Updated function for cleaning dates
import re
from datetime import datetime

def clean_dates(row):
    """
    Cleans and formats date-related information.

    Parameters:
    - row (pd.Series): A single row from the DataFrame.

    Returns:
    - tuple: A tuple containing the cleaned DATE and YEAR values.
    """
    title = row.get('TITLE', '')  # Use ES..TITLE or TITLE
    date_value = row.get('DATE', '')  # Use ES..DATE or DATE
    year_value = row.get('YEAR', '')  # Use ES..YEAR or YEAR

    date_str = date_value
    year_str = year_value

    if isinstance(title, str):
        date_pattern = r'(\b\w+\b) (\d{1,2}), (\d{4})'
        match = re.search(date_pattern, title, re.IGNORECASE)
        if match:
            month, day, year = match.groups()
            month_mapping = {
                'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04',
                'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08',
                'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12'
            }
            month_number = month_mapping.get(month.lower(), '01')
            date_str = f"{year}-{month_number}-{day.zfill(2)}"
            year_str = year

    return pd.Series([date_str, year_str])





def safe_extract_year(date_value):
    print(f"Raw value: {date_value}")  # Debugging: Log the raw input value
    try:
        # If the value is a number or string representation of a number, treat it as a year
        if isinstance(date_value, (int, float)) or (isinstance(date_value, str) and date_value.isdigit()):
            return int(date_value)

        # Try to parse as a complete date string
        date_obj = pd.to_datetime(date_value, errors='coerce')
        if pd.isnull(date_obj):
            print(f"Invalid date format: {date_value}")  # Debugging: Log invalid formats
            return np.nan

        year = date_obj.year
        print(f"Extracted year: {year} from {date_obj}")  # Debugging: Log extracted year
        return year
    except Exception as e:
        print(f"Error processing '{date_value}': {e}")  # Debugging: Log exceptions
        return np.nan

def extract_year_from_date(col):
    col = ensure_series(col)  # Ensure it's a Pandas Series
    return col.apply(safe_extract_year)





def clean_salutation(col):
    col = ensure_series(col).astype(str).fillna('')
    return col.str.strip().replace('no data', '')




# Function to clean the FROM and TO columns
def clean_from_to(col):
    col = ensure_series(col).astype(str).fillna('')  # Ensure it's a Pandas Series
    def format_from_to(text):
        words = text.split()
        if len(words) >= 2:
            first_name = words[0]
            last_name = words[-1]
            # Check if this name exists in the proper names dictionary
            if (first_name, last_name) in proper_names_list:
                return proper_names_list[(first_name, last_name)]
        return text  # Return the original if no match is found
    
    return col.apply(format_from_to)




def clean_relationship(col):
    col = ensure_series(col).astype(str).fillna('')
    return col.str.strip().replace('no data', '')




def clean_extent(col, language="English"):
    col = ensure_series(col).astype(str).fillna('')
    if language == "Spanish":
        return col.str.replace('hoja(s)', 'hojas').replace('páginas', 'página')
    else:
        return col.str.replace('leaf', 'leaves').replace('page(s)', 'pages')




def clean_notes(col):
    col = ensure_series(col).astype(str).fillna('')
    return col.str.strip().replace('no data', '').replace('; ', '[|]')




def clean_city_state_country_geoloc_english(df):
    """
    Cleans the state, country, and geolocation columns for English columns.
    Applies 'Unknown' for any missing data in the relevant columns.
    """
    # Ensure 'df' is a DataFrame
    if isinstance(df, pd.DataFrame):
        for column in ['SENDERS_STATE', 'SENDERS_COUNTRY', 'ADDRESSEES_STATE', 'ADDRESSEES_COUNTRY']:
            if column in df.columns:
                df[column] = df[column].apply(lambda x: "Unknown" if pd.isnull(x) else x)
    return df




def clean_city_state_country_geoloc_spanish(df):
    """
    Cleans the state, country, and geolocation columns for Spanish columns.
    Applies 'Desconocido' for any missing data in the relevant columns.
    """
    # Ensure 'df' is a DataFrame
    if isinstance(df, pd.DataFrame):
        for column in ['ES..SENDERS_STATE', 'ES..SENDERS_COUNTRY', 'ES..ADDRESSEES_STATE', 'ES..ADDRESSEES_COUNTRY']:
            if column in df.columns:
                df[column] = df[column].apply(lambda x: "Desconocido" if pd.isnull(x) else x)
    return df





def clean_coordinates(col):
    col = ensure_series(col).astype(str).fillna('')
    return col.str.replace('; ', '[|]').str.strip()





def fill_constant_values(df):
    # Define constants for both Spanish and English columns
    constant_values = {
        'DIGITAL_PUBLISHER': {
            'ES': 'Biblioteca de la Universidad Estatal de Nuevo México',
            'EN': 'New Mexico State University Library'
        },
        'SOURCE': {
            'ES': 'Archivos y colecciones especiales de la biblioteca de NMSU',
            'EN': 'NMSU Library Archives and Special Collections'
        },
        'UNIT': {
            'ES': 'Colecciones históricas de Río Grande',
            'EN': 'Rio Grande Historical Collections'
        },
        'FORMAT': {
            'ES': 'la aplicación/pdf',
            'EN': 'application/pdf'
        },
        'TYPE': {
            'ES': 'Texto',
            'EN': 'Text'
        },
        'ACCESS_RIGHTS': {
            'ES': 'Abierto para la reutilización',
            'EN': 'Open for re-use'
        },
        # Non-language specific constants
        'OA_NAME': '',
        'ES..OA_DESCRIPTION': 'Esta colección está disponible en inglés y español',
        'OA_DESCRIPTION': 'This collection is available in both, English and Spanish',
        'OA_COLLECTION': '10317',
        'OA_PROFILE': 'Documents',
        'OA_STATUS': 'PUBLISH',
        'OA_LINK': '',
        'OA_LOG': '',
        'OA_OBJECT_TYPE': 'RECORD',
        'OA_METADATA_SCHEMA': '4',
        'OA_FEATURED': '0'
    }

    # For each column, create or overwrite with the constant values
    for column_name, value in constant_values.items():
        # Determine the number of rows to fill
        num_rows = len(df)
        
        if isinstance(value, dict):
            # Handle language-specific columns
            df[f'ES..{column_name}'] = pd.Series([value['ES']] * num_rows)
            df[column_name] = pd.Series([value['EN']] * num_rows)
        else:
            # For single-value columns
            df[column_name] = pd.Series([value] * num_rows)

    return df

def clean_columns_in_sheets(xls, column_cleaning_rules):
    """
    Cleans all columns in all sheets of the provided Excel file, applying
    the specified cleaning rules and transformations.

    Parameters:
    - xls (pd.ExcelFile): The loaded Excel file with multiple sheets.
    - column_cleaning_rules (dict): Dictionary mapping column names to cleaning functions.

    Returns:
    - None: Saves the cleaned DataFrame for each sheet to the global `cleaned_sheets` dictionary.
    """
    for sheet_name in xls.sheet_names:
        # Read the sheet into a DataFrame
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Skip specific sheet (e.g., technical metadata sheet)
        if sheet_name == "GO_Technical metadata":
            print(f"Skipping sheet: {sheet_name}")
            continue

        print(f"Processing sheet: {sheet_name}")

        # Store the original column order to preserve it later
        original_columns = df.columns.tolist()

        # Fill missing values and replace "no data" placeholders
        df = df.fillna('').replace('no data', '')

       # Apply column-specific cleaning rules
   
        # Apply column-specific cleaning rules
    
    for column_name, cleaning_func in column_cleaning_rules.items():
        if column_name in df.columns and 'DATE' in column_name:
            # Ensure the YEAR column exists
            year_column = f"{column_name}_YEAR"
            if year_column not in df.columns:
                df[year_column] = ""

            # Apply the cleaning function to the entire row
            df[[column_name, year_column]] = df.apply(
                lambda row: pd.Series(cleaning_func(row)),
                axis=1
            )
        else:
            # Apply cleaning function column-wise for other columns
            df[column_name] = df[column_name].apply(cleaning_func)



            # Add DIGITAL_IDENTIFIER transformation explicitly
        for col_name in ["DIGITAL_IDENTIFIER", "ES..DIGITAL_IDENTIFIER"]:
                if col_name in df.columns:
                    print(f"Cleaning column: {col_name}")
                    df[col_name] = clean_digital_identifier(df, col_name)


            # Reorder columns to match the original column order
                df = df[original_columns]

                        # Fill constant values for specific columns
                df = fill_constant_values(df)

                        # Save the cleaned DataFrame into a global dictionary for later use
                cleaned_sheets[sheet_name] = df



def debug_cleaning_func(row, cleaning_func):
    result = cleaning_func(row)
    print(f"Row: {row}, Result: {result}")  # Debug output
    return result





# Mapping of columns to cleaning functions (exactly as you provided)
column_cleaning_rules = {

    # FullFolderFilePath columns
    "FullFolderFilePath": lambda col: create_full_folder_file_path(col),



    # TITLE columns
    "ES..TITLE": lambda col: clean_title(col, language="Spanish"),
    "TITLE": clean_title_english,
    
    
    # SERIES columns
    "ES..SERIES": clean_series,
    "SERIES": clean_series,
    
    # BOX_FOLDER columns
    "ES..BOX_FOLDER": clean_box_folder,
    "BOX_FOLDER": clean_box_folder,
    
    # COLLECTION_NAME columns
    "ES..COLLECTION_NAME": lambda col: clean_collection_name(col, language="Spanish"),
    "COLLECTION_NAME": lambda col: clean_collection_name(col, language="English"),
    
    # COLLECTION_NUMBER columns
    "ES..COLLECTION_NUMBER": clean_collection_number,
    "COLLECTION_NUMBER": clean_collection_number,

    # DATE columns
    "ES..DATE": lambda df: clean_dates(df, 'ES..TITLE', 'ES..DATE', 'ES..YEAR'),
    "DATE": lambda df: clean_dates(df, 'TITLE', 'DATE', 'YEAR'),

    
    # YEAR columns
    "ES..YEAR": extract_year_from_date,
    "YEAR": extract_year_from_date,
    
    # SALUTATION columns
    "ES..SALUTATION": clean_salutation,
    "SALUTATION": clean_salutation,
    
    # FROM columns
    "ES..FROM": clean_from_to,
    "FROM": clean_from_to,
    
    # TO columns
    "ES..TO": clean_from_to,
    "TO": clean_from_to,
    
    # RELATIONSHIP1 columns
    "ES..RELATIONSHIP1": clean_relationship,
    "RELATIONSHIP1": clean_relationship,
    
    # RELATIONSHIP2 columns
    "ES..RELATIONSHIP2": clean_relationship,
    "RELATIONSHIP2": clean_relationship,
    
    # EXTENT columns
    "ES..EXTENT": lambda col: clean_extent(col, language="Spanish"),
    "EXTENT": lambda col: clean_extent(col, language="English"),
    
    # NOTES columns
    "ES..NOTES": clean_notes,
    "NOTES": clean_notes,

    # Combined CITY, STATE, COUNTRY columns
    # English CITY, STATE, COUNTRY columns
    "SENDERS_CITY": clean_city_state_country_geoloc_english,
    "SENDERS_STATE": clean_city_state_country_geoloc_english,
    "SENDERS_COUNTRY": clean_city_state_country_geoloc_english,
    "ADDRESSEES_CITY": clean_city_state_country_geoloc_english,
    "ADDRESSEES_STATE": clean_city_state_country_geoloc_english,
    "ADDRESSEES_COUNTRY": clean_city_state_country_geoloc_english,

    # Spanish CITY, STATE, COUNTRY columns
    "ES..SENDERS_CITY": clean_city_state_country_geoloc_spanish,
    "ES..SENDERS_STATE": clean_city_state_country_geoloc_spanish,
    "ES..SENDERS_COUNTRY": clean_city_state_country_geoloc_spanish,
    "ES..ADDRESSEES_CITY": clean_city_state_country_geoloc_spanish,
    "ES..ADDRESSEES_STATE": clean_city_state_country_geoloc_spanish,
    "ES..ADDRESSEES_COUNTRY": clean_city_state_country_geoloc_spanish,

    # COORDINATES columns
    "ES..GEOLOC_SCITY": clean_coordinates,
    "GEOLOC_SCITY": clean_coordinates,
    
    # POST_SCRIPTUM columns
    "ES..POST_SCRIPTUM": clean_notes,
    "POST_SCRIPTUM": clean_notes,
    
    # SIGNATURE columns
    "ES..SIGNATURE": clean_from_to,
    "SIGNATURE": clean_from_to,
    
    # PHYSICAL_DESCRIPTION columns
    "ES..PHYSICAL_DESCRIPTION": clean_notes,
    "PHYSICAL_DESCRIPTION": clean_notes,
    
    # OTHER_PEOPLE_MENTIONED columns
    "ES..OTHER_PEOPLE_MENTIONED": clean_notes,
    "OTHER_PEOPLE_MENTIONED": clean_notes,
    
    # OTHER_PLACES_MENTIONED columns
    "ES..OTHER_PLACES_MENTIONED": clean_notes,
    "OTHER_PLACES_MENTIONED": clean_notes,
    
    # SUBJECT_LCSH columns
    "ES..SUBJECT_LCSH": clean_notes,
    "SUBJECT_LCSH": clean_notes,
}

# Apply cleaning to all sheets at once
clean_columns_in_sheets(xls, column_cleaning_rules)

# Create transformed output file
output_file_path = os.path.join(os.getcwd(), f"Transformed_{file_name}")

# Save the cleaned data to a new Excel file
with pd.ExcelWriter(output_file_path) as writer:
    for sheet_name, cleaned_df in cleaned_sheets.items():
        cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Data cleaned and saved to: {output_file_path}")