import pandas as pd

# Load the two files
original_df = pd.read_excel('Test2.xlsx')
transformed_df = pd.read_excel('Transformed_Test2.xlsx')

# Print column name differences for debugging
print("Columns in Original DataFrame but not in Transformed DataFrame:", set(original_df.columns) - set(transformed_df.columns))
print("Columns in Transformed DataFrame but not in Original DataFrame:", set(transformed_df.columns) - set(original_df.columns))

# Align DataFrames by their common columns
common_columns = original_df.columns.intersection(transformed_df.columns)
original_df = original_df[common_columns]
transformed_df = transformed_df[common_columns]

# Reset index to ensure they are aligned
original_df = original_df.reset_index(drop=True)
transformed_df = transformed_df.reset_index(drop=True)

# Now compare the DataFrames
comparison_original_vs_transformed = original_df.compare(transformed_df)

# Output the comparison to see differences
print(comparison_original_vs_transformed)
