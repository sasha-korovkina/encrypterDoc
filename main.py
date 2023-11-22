import aspose.pdf as ap
import os
import re
import pandas as pd
import random
import string

# variables
input_path = "M:\CDB\Analyst\Rhys\Python\CustodianExtract\custodian_extraction\input\Barclays Capital Sec\@Barclays Capital Sec Uk@ BE0003818359 31102023.pdf"
output_path = "M:\CDB\Analyst\Rhys\Python\CustodianExtract\custodian_extraction\input\Barclays Capital Sec\Barclays Capital Sec Uk.pdf"
directory_path = r"M:\CDB\Analyst\Rhys\Python\CustodianExtract\custodian_extraction\input\Barclays Capital Sec"

old_word = "BE0003818359"
new_word = "US5949181045"

def extract_isin_codes(directory_path):
    # Get all files in the directory
    files = [f for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f))]

    # Define a regular expression pattern to match 12-letter number codes
    pattern = re.compile(r'\b[A-Za-z]{2}[0-9]{10}\b')

    # Extract ISIN codes from file names
    isin_codes = []
    for file in files:
        match = pattern.search(file)
        if match:
            isin_codes.append(match.group())

    return isin_codes

def generate_random_sequence(isin_format):
    letters = [random.choice(string.ascii_uppercase) for _ in range(2)]
    digits = [random.choice(string.digits) for _ in range(10)]
    return ''.join(letters + digits)

def wordReplace(input_path, output_path, old_word, new_word):
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

    # Load the document
    doc = ap.Document(input_path)

    # Create a text absorber
    txtAbsorber = ap.text.TextFragmentAbsorber(old_word)

    # Search text
    doc.pages.accept(txtAbsorber)

    # Get reference to the found text fragments
    textFragmentCollection = txtAbsorber.text_fragments

    # Parse all the searched text fragments and replace text
    for txtFragment in textFragmentCollection:
        txtFragment.text = new_word

    doc.save(output_path)

wordReplace(input_path, output_path, old_word, new_word)
