'''
Algo:
1. The file goes threough the folder and gets the isin which will be replaced in each file
2. The file is opened and the isin is replaced
3.
'''
import aspose.pdf as ap
import os
import re
import pandas as pd
import random
import string
import shutil

directory_path = r"M:\CDB\Analyst\Rhys\Python\CustodianExtract\custodian_extraction\input\Barclays Capital Sec"
output_directory_path = r"M:\CDB\Analyst\Rhys\Python\CustodianExtract\custodian_extraction\output\Barclays Capital Sec"

def extract_isin_codes(filename):
    # Define a regular expression pattern to match 12-letter number codes
    pattern = re.compile(r'\b[A-Za-z]{2}[0-9]{10}\b')

    match = pattern.search(filename)
    if match:
        isin = match.group()
        print(f"Found ISIN in the defined format: {isin}")
        return isin
    else:
        print("No ISIN found in the defined format.")
        return None

def generate_random_sequence(isin_format):
    letters = [random.choice(string.ascii_uppercase) for _ in range(2)]
    digits = [random.choice(string.digits) for _ in range(10)]
    return ''.join(letters + digits)

def isin_dataframe():
    isin_codes = extract_isin_codes(directory_path)
    for isin_code in isin_codes:
        print(isin_code)

    random_sequences = [generate_random_sequence(isin_format) for isin_format in isin_codes]
    # Create a DataFrame
    df = pd.DataFrame({
        'ISIN': isin_codes,
        'RandomSequence': random_sequences
    })

    # Print the DataFrame
    print(df)


def wordReplace(directory_path):
    print("Files in the directory:")

    for filename in os.listdir(directory_path):
        full_path = os.path.join(directory_path, filename)
        print(full_path)

        # Load the document
        doc = ap.Document(full_path)

        # Extract the ISIN from the title
        old_word = extract_isin_codes(filename)

        # Create a text absorber
        txtAbsorber = ap.text.TextFragmentAbsorber(old_word)

        # Search text
        doc.pages.accept(txtAbsorber)

        # Get reference to the found text fragments
        textFragmentCollection = txtAbsorber.text_fragments

        # Generate a new ISIN
        new_word = generate_random_sequence(old_word)
        print(f"New ISIN: {new_word}")

        # Replace text in the document
        for txtFragment in textFragmentCollection:
            txtFragment.text = new_word

        # Create the new filename with the prefix "ENCRYPTED" and keep the PDF extension
        new_filename = f"ENCRYPTED_{new_word}.pdf"

        # Save the file at the same path with the new filename
        save_path = os.path.join(os.path.dirname(full_path), new_filename)
        print(f"Saving file at: {save_path}")
        #doc.save(save_path)
        new_full_path = os.path.join(output_directory_path, os.path.splitext(filename)[0] + ".xls")
        #output_new = os.path.join(output_directory_path, os.path.splitext(filename)[0]  + new_word + ".xls")
        output_new = os.path.join(output_directory_path, f"ENCRYPTED_{new_word}.xlsx")

        if os.path.exists(new_full_path):
            shutil.copy2(new_full_path, output_new)
            print(f"File copied to: {output_new}")
        else:
            print(f"File not found at: {new_full_path}")
        print(output_new)


wordReplace(directory_path)
