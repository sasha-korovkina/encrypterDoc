from docx import Document
import random
import string

def find_and_replace(doc, old_text, new_text):
    for paragraph in doc.paragraphs:
        if old_text in paragraph.text:
            paragraph.text = paragraph.text.replace(old_text, new_text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if old_text in cell.text:
                    cell.text = cell.text.replace(old_text, new_text)

def generate_random_sequence(isin_format):
    letters = [random.choice(string.ascii_uppercase) for _ in range(2)]
    digits = [random.choice(string.digits) for _ in range(10)]
    return ''.join(letters + digits)

# Load the Word document
doc_path = "M:\CDB\Analyst\Rhys\Python\CustodianExtract\custodian_extraction\input\BNP Paribase Singapore\@BNP Paribas SS Singapore@ LU1598757687 02102023.docx"
doc = Document(doc_path)

# Specify the word to find and replace
old_word = 'LU1598757687'
new_word = generate_random_sequence(old_word)

# Call the find_and_replace function
find_and_replace(doc, old_word, new_word)

# Save the modified document
new_filename = f"ENCRYPTED_{new_word}.docx"
doc.save("M:\\CDB\\Analyst\\Rhys\\Python\\CustodianExtract\\custodian_extraction\\input\\BNP Paribase Singapore\\" + new_filename)
