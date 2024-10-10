import pandas as pd
import numpy as np
import re
import os
import argparse
import unicodedata 


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





# Ensure the file exists in the directory
if not os.path.exists(file_path):
    raise FileNotFoundError(f"The file {file_name} does not exist in the current directory.")




# Load spreadsheet with multiple sheets
xls = pd.ExcelFile(file_path)





# Dictionary to hold cleaned DataFrames
cleaned_sheets = {}



# Load the known proper names from the Excel file
proper_names_df = pd.read_excel('ProperNames.xlsx')  # Adjust file path as needed


proper_names_list = proper_names_df['PROPER AUTHORIZED NAMES'].str.strip().tolist()  # Create a list of proper names





def ensure_series(col):
    """Ensures the input is a Pandas Series."""
    if isinstance(col, pd.Series):
        return col
    elif isinstance(col, list) or isinstance(col, np.ndarray):
        return pd.Series(col)
    else:
        return pd.Series([col])

# Define column-specific cleaning functions




def clean_digital_identifier(col):
    col = ensure_series(col).astype(str).fillna('')  
    def clean_value(x):
        if pd.isnull(x):
            return x
        x = str(x).strip().lower()
        x = x.replace("ms0004", "Ms0004")  # Ensure the "Ms" is capitalized
        x = x.replace("mS0004", "Ms0004")
        return x if x.endswith('.pdf') else f"{x}.pdf"
    
    return col.apply(clean_value)





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
    """
    Normalize Spanish accents for proper display of words like 'invitación'
    """
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


def clean_dates(row, title_column, date_column, year_column):
    title = row[title_column]

    # Ensure title is a string and handle None or NaN values
    title_str = str(title) if title else ''

    # Debugging: Output the type and value for each row's title
    print(f"Row {row.name} - Title value: {title_str}")

    # Pattern to match dates like "Marzo 7, 1924"
    date_pattern = r'(\b\w+\b) (\d{1,2}), (\d{4})'
    match = re.search(date_pattern, title_str, re.IGNORECASE)

    if match:
        month, day, year = match.groups()

        # Map Spanish months to their English equivalents or month numbers
        month_mapping = {
            'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04',
            'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08',
            'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12'
        }

        # Convert month name to month number
        month_number = month_mapping.get(month.lower(), '01')  # Default to January if not found

        # Construct a formatted date string
        formatted_date = f'{year}-{month_number}-{day.zfill(2)}'
        row[date_column] = formatted_date
        row[year_column] = year  # Store the year as well
    elif 'sin fecha' in title_str.lower() or 'undated' in title_str.lower():
        # Clear date and year for "sin fecha" or "undated"
        row[date_column] = ''
        row[year_column] = ''
    else:
        # No valid date found, clear both date and year
        row[date_column] = ''
        row[year_column] = ''

    return row



def extract_year_from_date(col):
    col = ensure_series(col)
    return pd.to_datetime(col, errors='coerce').dt.year




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
            'ES': 'la applicación/pdf',
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
        'OA_NAME': {
            'ES': '',
            'EN': ''
        },
        'OA_DESCRIPTION': {
            'ES': 'Esta colleción está disponible en inglés y español',
            'EN': 'This collection is available in both, English and Spanish'
        },
        'OA_COLLECTION': {
              '10317'
            
        },
        'OA_PROFILE': {
            'Documents'
        },
        'OA_STATUS': {
            'PUBLISH'
        },
        'OA_LINK': {
            ''
        },
        'OA_LOG': (
            ''
        ),
        'OA_OBJECT_TYPE': {
            'RECORD'
        },
        'OA_METADATA_SCHEMA': {
            '4'
        },
        'OA_FEATURED': {
            '0'
        }
    }



def clean_columns_in_sheets(xls, column_cleaning_rules):
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Ignore the GO_Technical metadata sheet
        if sheet_name == "GO_Technical metadata":
            print(f"Skipping sheet: {sheet_name}")
            continue
        
        print(f"Processing sheet: {sheet_name}")
        
        # Fill any missing values with blanks and replace "no data" with blanks
        df = df.fillna('').replace('no data', '')
        
        # Apply cleaning functions
        for column_name, cleaning_func in column_cleaning_rules.items():
            if column_name in df.columns:
                print(f"Applying cleaning to column: {column_name}")
            try:
                
                # If column is related to 'DATE', apply row-wise cleaning
                if 'DATE' in column_name:
                    # Debugging: inspect what the cleaning function returns
                    date_col = column_name
                    year_col = 'YEAR' if 'ES..' not in column_name else 'ES..YEAR'
                    df[[date_col, year_col]] = df.apply(lambda row: pd.Series(cleaning_func(row, 'TITLE', date_col, year_col)), axis=1)
                else:
                    # Apply column-wise cleaning
                    df[column_name] = df[column_name].apply(cleaning_func)
        
            except Exception as e: 
                print(f"Error applying {cleaning_func.__name__} to {column_name}: {e}")
                    
        # Save cleaned DataFrame
        cleaned_sheets[sheet_name] = df

def debug_cleaning_func(row, cleaning_func):
    result = cleaning_func(row)
    print(f"Row: {row}, Result: {result}")  # Debug output
    return result





# Mapping of columns to cleaning functions (exactly as you provided)
column_cleaning_rules = {

    # FullFolderFilePath columns
    "FullFolderFilePath": lambda col: create_full_folder_file_path(col),

    # DIGITAL_IDENTIFIER columns
    "DIGITAL_IDENTIFIER": clean_digital_identifier,
    "ES..DIGITAL_IDENTIFIER": clean_digital_identifier,
    
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
