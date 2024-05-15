import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
from nltk.tag import pos_tag
import xlwings as xw

nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
# print(nltk.data.path)

# Define the path to your password-protected Excel file
file_path = '/Users/hhz240/Documents/Research Culture PGR project/data/PGR Supervisor Clean Data V1.xlsx'
sheet_name = 'PT'  # Specify the sheet name containing your data
excel_password = 'PeeGeeArr2024!'  # Replace with your actual password

# Configure xlwings app settings to suppress Excel window
app = xw.App(visible=False, add_book=False)  # Create a hidden Excel instance

# Load your dataset into a pandas DataFrame
try:
    # app = xw.App(visible=False)  # Create a hidden Excel instance
    # wb = xw.Book(file_path, password=excel_password)  # Open the Excel file with password
    wb = app.books.open(file_path, password=excel_password, update_links=False)  # Open the Excel file with password
    sheet = wb.sheets[sheet_name]  # Access the specified sheet

    # Determine the range of the table and load into DataFrame
    data_range = sheet.used_range  # Get the used range of the sheet
    data_values = data_range.value  # Get values from the used range

    # Find the index of the row containing the desired headers
    header_index = None
    for idx, row in enumerate(data_values):
        if "Student ID" in row:
            header_index = idx
            break

    if header_index is not None:
        # Extract column headers from the identified row
        headers = data_values[header_index]
        
        # Create DataFrame from rows below the header row
        data = pd.DataFrame(data_values[header_index + 1:], columns=headers)
    else:
        raise ValueError("Header row not found in the specified sheet.")

except Exception as e:
    print(f"Error: Unable to load Excel file. {str(e)}")
    exit(1)  # Exit the script if there's an error loading the file
finally:
    app.display_alerts = False  # Disable display alerts (e.g., save changes prompt)
    app.screen_updating = False  # Disable screen updating (hide Excel window)
    app.quit()  # Quit Excel application


# print("Loaded DataFrame:")
# print(data.head(5))

 
# Tokenize and tag programme titles
def extract_disciplines(programme_title):
    tokens = word_tokenize(programme_title)
    tagged_tokens = pos_tag(tokens)
    disciplines = [token for token, pos in tagged_tokens if pos == 'NNP']  # Extract proper nouns (assumed as disciplines)
    return disciplines

# Apply extraction function to each programme title
data['Disciplines'] = data['Programme Title'].apply(extract_disciplines)

# Save the modified DataFrame to a CSV file
output_csv_path = '/Users/hhz240/Documents/Research Culture PGR project/data/PGR_disciplines.csv'  # Specify the desired output file path
data.to_csv(output_csv_path, index=False)  # Save DataFrame to CSV without including row indices

print(data.head(5))


# Count unique combinations of disciplines
unique_discipline_combinations = data['Disciplines'].apply(tuple).nunique()

# Display the count of unique interdisciplinary programme titles
print(f"Number of research degrees with interdisciplinary elements: {unique_discipline_combinations}")
