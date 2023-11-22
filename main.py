from encyptor_pdf import wordReplace
import pandas as pd

# Company names generator
company_names = ["GALAPAGOS NV", "TECK RESOURCES LTD", "IMPLENIA AG", "BANCO SANTANDER REG. SHS", "BANCO SANTANDER ADR REG. SHS", "KONE OYJ -B- REG. SHS", "DZ PRIVATBANK S.A."]
random_names = ["AMAZON.COM INC", "FACEBOOK INC.", "JOHNSON & JOHNSON","ALPHABET INC.", "APPLE INC.", "GOOGLE INC.", "MICROSOFT CORP."]

df = pd.DataFrame({
    'Company': company_names,
    'Random_Name': random_names
})
print(df)

directory_path = r"M:\CDB\Analyst\Rhys\Python\CustodianExtract\custodian_extraction\input\Barclays Capital Sec"
output_directory_path = r"M:\CDB\Analyst\Rhys\Python\CustodianExtract\custodian_extraction\output\Barclays Capital Sec"
wordReplace(directory_path, df, output_directory_path)
