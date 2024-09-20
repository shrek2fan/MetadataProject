import pandas as pd
from difflib import get_close_matches
import sys
import re

# Function to clean and standardize names
def clean_name(name):
    return name.strip().title()  # Strip whitespace and capitalize each word

# Function to apply title casing and handle connectors (e.g., "a" to "to")
def format_title(title, is_spanish=False):
    # Strip leading/trailing spaces
    title = title.strip()
    
    # Handle connectors (assume Spanish uses 'a', English uses 'to')
    if is_spanish:
        title = re.sub(r'\bto\b', 'a', title, flags=re.IGNORECASE)  # Ensure Spanish uses 'a'
    else:
        title = re.sub(r'\ba\b', 'to', title, flags=re.IGNORECASE)  # Ensure English uses 'to'

    # Capitalize proper nouns, and ensure the first letter is capitalized
    return title.capitalize()

# Function to find the closest match for a name in the authorized names list
def find_closest_authorized_name(name, authorized_names):
    # Use difflib to find the closest match from the authorized list
    matches = get_close_matches(name, authorized_names, n=1, cutoff=0.8)
    return matches[0] if matches else None

# Load the authorized names CSV file (this remains constant)
authorized_names_df = pd.read_csv('Controlled_vocabularies_NEH_Amador_(PEOPLE).csv', encoding='ISO-8859-1')

# Clean the authorized names list by removing any unnecessary rows and focusing on the 'PEOPLE' column
authorized_names_cleaned_df = authorized_names_df.dropna(subset=['PEOPLE'])
authorized_names_list = authorized_names_cleaned_df['PEOPLE'].tolist()

# Get the input metadata file from the user (via command-line argument or prompt)
if len(sys.argv) > 1:
    metadata_file = sys.argv[1]  # Take the first command-line argument as the file name
else:
    metadata_file = input("Please enter the path to the metadata Excel file (.xlsx): ")

# Load the provided metadata file (assuming 'OA_Descriptive metadata' sheet is relevant)
try:
    metadata_descriptive_df = pd.read_excel(metadata_file, sheet_name='OA_Descriptive metadata')
except Exception as e:
    print(f"Error loading file: {e}")
    sys.exit(1)

# Apply transformations to 'FROM', 'TO', 'ES..TITLE', and 'TITLE' columns
for index, row in metadata_descriptive_df.iterrows():
    # Clean and standardize names
    es_from_clean = clean_name(row['ES..FROM']) if pd.notnull(row['ES..FROM']) else ''
    en_from_clean = clean_name(row['FROM']) if pd.notnull(row['FROM']) else ''
    
    # Find closest matches in authorized names list
    es_from_corrected = find_closest_authorized_name(es_from_clean, authorized_names_list)
    en_from_corrected = find_closest_authorized_name(en_from_clean, authorized_names_list)
    
    # Replace with corrected names if a close match is found
    if es_from_corrected:
        metadata_descriptive_df.at[index, 'ES..FROM'] = es_from_corrected
    if en_from_corrected:
        metadata_descriptive_df.at[index, 'FROM'] = en_from_corrected
    
    # Apply title formatting to 'ES..TITLE' and 'TITLE' (handle language-based connectors)
    es_title_clean = format_title(row['ES..TITLE'], is_spanish=True) if pd.notnull(row['ES..TITLE']) else ''
    en_title_clean = format_title(row['TITLE'], is_spanish=False) if pd.notnull(row['TITLE']) else ''
    
    metadata_descriptive_df.at[index, 'ES..TITLE'] = es_title_clean
    metadata_descriptive_df.at[index, 'TITLE'] = en_title_clean

# Save the modified metadata to a new Excel file for review
output_file = f"Transformed_{metadata_file.split('/')[-1]}"
metadata_descriptive_df.to_excel(output_file, index=False)

print(f"Transformation completed. Please check '{output_file}' for results.")
