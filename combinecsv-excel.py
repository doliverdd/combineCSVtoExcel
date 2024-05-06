import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment

# Directory containing CSV files..PUT YOUR CSV FILES HERE
directory = os.getcwd() + '/files'

# List CSV files in the directory
csv_files = [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith('.csv')]

# Combine CSV files into a single DataFrame
combined_df = pd.concat([pd.read_csv(file) for file in csv_files], ignore_index=True)

# Remove specific text string from column E
combined_df['ChatTranscript'] = combined_df['ChatTranscript'].str.replace("Bot says: Hello I'm ITSM a virtual assistant. Just so you are aware "
"I sometimes use AI to answer your questions. If you provided a website during creation try asking me about it! Next try giving me some more"
"knowledge by setting up generative AI.;User says: What software is available to students;Bot says: Please wait a moment while we look through "
"our knowledge base;", '')

# Write DataFrame to Excel
excel_file = 'files/combined_output.xlsx'
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    combined_df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Get the Excel workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Set column widths to approximately 5 inches (assuming 72 pixels per inch)
    for i, column in enumerate(combined_df.columns):
        column_width = 5 * 72  # 5 inches converted to pixels
        worksheet.column_dimensions[chr(65 + i)].width = column_width / 7  # divide by 7 for a good fit

        # Set row height to 2 inches
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)
            worksheet.row_dimensions[row[0].row].height = 2 * 72  # 2 inches converted to pixels

    # Wrap text in each column
    # for row in worksheet.iter_rows(min_row=1, max_row=1):
    # for cell in row:
           # cell.alignment = Alignment(wrap_text=True)

print("Excel file with combined data and formatted column widths has been created: combined_output.xlsx")