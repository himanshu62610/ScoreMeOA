import pdfplumber
import pandas as pd
import os

def group_words_to_lines(words, vertical_threshold=3):
    """
    Group words into lines based on their 'top' coordinate.
    """
    words.sort(key=lambda w: (w['top'], w['x0']))
    lines, current_line = [], []
    current_top = None
    
    for word in words:
        if current_top is None or abs(word['top'] - current_top) <= vertical_threshold:
            current_line.append(word)
        else:
            lines.append(current_line)
            current_line = [word]
        current_top = word['top']
    
    if current_line:
        lines.append(current_line)
    return lines

def get_columns_from_line(text_items, gap_threshold=5):
    """
    Group words into columns based on horizontal spacing.
    """
    text_items.sort(key=lambda item: item['x0'])
    columns, current_column = [], []
    previous_x1 = None
    
    for item in text_items:
        if previous_x1 is not None and (item['x0'] - previous_x1) > gap_threshold:
            columns.append(" ".join(current_column))
            current_column = []
        current_column.append(item['text'])
        previous_x1 = item['x1']
    
    if current_column:
        columns.append(" ".join(current_column))
    return columns

def extract_pdf_table_structure(pdf_path):
    """
    Extract text from a PDF and organize it into a table-like structure.
    """
    table_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            lines = group_words_to_lines(words)
            for line in lines:
                columns = get_columns_from_line(line)
                table_data.append(columns)
    return table_data

def save_table_data_to_excel(table_data, excel_path):
    """
    Save extracted table data to an Excel file.
    """
    df = pd.DataFrame(table_data)
    df.to_excel(excel_path, index=False, header=False)
    print(f"Excel saved to: {excel_path}")

def process_pdfs_in_folder(pdf_folder, output_folder):
    """
    Process all PDFs in a folder and save extracted tables as Excel files.
    """
    os.makedirs(output_folder, exist_ok=True)
    
    for file in os.listdir(pdf_folder):
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, file)
            output_excel = os.path.join(output_folder, f"{os.path.splitext(file)[0]}.xlsx")
            print(f"Processing: {file}")
            table_data = extract_pdf_table_structure(pdf_path)
            save_table_data_to_excel(table_data, output_excel)

if __name__ == "__main__":
    pdf_folder = "Input_pdfs"  # Folder containing PDF files
    output_folder = "extracted_tables"  # Folder to save Excel files
    process_pdfs_in_folder(pdf_folder, output_folder)
