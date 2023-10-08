import os
import shutil
import time
from datetime import date, datetime
import hashlib
import json
import re
import tabula

import pandas as pd
from currency_converter import CurrencyConverter, ECB_URL

# Create a single instance of CurrencyConverter
c = CurrencyConverter(ECB_URL, fallback_on_wrong_date=True, fallback_on_missing_rate=True)

### UTILITY FUNCTIONS

# Define a function to calculate the exchange rate
def calculate_target_currency_rate(row, c):
    try:
        specified_date = row['Date']
        extract_date = datetime.strptime(specified_date, '%d-%m-%Y')
        this_date = date(extract_date.year, extract_date.month, extract_date.day)
        base_currency = row['Currency']  # Replace 'Currency' with the actual column name for currency
        target_currency = "DKK"  # You specified DKK as the target currency
        amount = row['Amount_currency']  # Replace 'Amount_Currency' with the actual column name for the amount in currency

        # Perform the currency conversion with error handling
        converted_amount = c.convert(amount, base_currency, target_currency, this_date)

        return round(converted_amount, 2)

    except Exception as e:
        return None  # Handle errors by returning None or any suitable value

def get_currency_rate(row, c):
    try:
        specified_date = row['Date']
        extract_date = datetime.strptime(specified_date, '%d-%m-%Y')
        this_date = date(extract_date.year, extract_date.month, extract_date.day)
        base_currency = row['Currency']  # Replace 'Currency' with the actual column name for currency
        target_currency = "DKK"  # You specified DKK as the target currency

        # Perform the currency conversion with error handling
        currency_rate = c.convert(1.00, base_currency, target_currency, this_date)

        return round(currency_rate, 2)

    except Exception as e:
        return None  # Handle errors by returning None or any suitable value

def is_file_open(file_path):
    try: 
        if os.path.isfile(file_path):
            os.rename(file_path, file_path)
        return False
    except OSError:
        return True

def backup_csv_file(csv_file_path):
    try:
        # Check if the CSV file exists
        if not os.path.isfile(csv_file_path):
            print(f"CSV file '{csv_file_path}' does not exist. Skipping backup.")
            return

        # Check if the CSV file is empty
        if os.path.getsize(csv_file_path) == 0:
            print(f"CSV file '{csv_file_path}' is empty. Skipping backup.")
            return

        # Create a backup file name by adding a timestamp to the original file name
        backup_file_path = csv_file_path.replace('.csv', f'_backup_{time.strftime("%Y%m%d%H%M%S")}.csv')

        # Copy the original CSV file to the backup location
        shutil.copyfile(csv_file_path, backup_file_path)

        print(f"Backup created: {backup_file_path}")
    except Exception as e:
        print(f"An error occurred while creating a backup: {str(e)}")

### DATA EXTRACTION FUNCTIONS

def extract_csv_data(file_path, bank):
    try:
        print(f"Extracting data from CSV file: {file_path}")
        # Specific settings for Danske Bank
        if bank == 'Danske Bank':
            delimiter = ';'
            encoding = 'ISO-8859-1'

            # Read the CSV file into a pandas DataFrame with Danske Bank-specific settings
            df = pd.read_csv(file_path, delimiter=delimiter, encoding=encoding)

            # Rename columns for consistency and readability for Danske Bank
            df = df.rename(columns={'Dato': 'Date', 'Tekst': 'Description', 'Beløb': 'Amount'})

            # Format the 'Date' column to 'DD-MM-YYYY' format for Danske Bank
            df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y').dt.strftime('%d-%m-%Y')

            # Remove thousand separator (dot) and replace comma with dot for decimal delimiter in 'Amount' column for Danske Bank
            df['Amount'] = df['Amount'].str.replace('.', '').str.replace(',', '.', regex=False).astype(float)

            # Create unique identifier
            df['UniqueID'] = df.apply(lambda row: hashlib.md5(f"{row['Date']}_{row['Description']}_{row['Amount']}_{row['Saldo']}".encode()).hexdigest(), axis=1)

            # Create bank identifier
            df['Bank'] = 'Danske Bank'
            df['Amount_currency'] = df['Amount']
            df['Currency'] = 'DKK'
            df['Currency_Rate'] = 1

            selected_columns = ['Date', 'Description', 'Amount', 'Amount_currency', 'Currency', 'Currency_Rate', 'Bank', 'UniqueID']
            df_selected = df[selected_columns]

        elif bank == 'Wise':
            # Specific settings for Wise bank (add more banks as needed)
            delimiter = ','  # Change this delimiter as per the CSV format
            encoding = 'utf-8'  # Change encoding if necessary

            # Read the CSV file into a pandas DataFrame with settings for Wise bank
            df = pd.read_csv(file_path, delimiter=delimiter, encoding=encoding)

            # Filter out unwanted rows
            df = df[(df['Status'] != 'CANCELLED') & (df['Target amount (after fees)'] != 0) & (df['Source amount (after fees)'] != 0)]

            # Format the date column to 'DD-MM-YYYY' format
            df['Date'] = pd.to_datetime(df['Finished on'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d-%m-%Y')

            # Initialize empty columns
            df['Description'] = ''
            df['Amount_currency'] = ''
            df['Currency'] = ''

            # Set values based on 'Direction' column
            df.loc[df['Direction'] == 'IN', 'Description'] = df['Source name']
            df.loc[df['Direction'] == 'OUT', 'Description'] = df['Target name']
            df.loc[df['Direction'] == 'NEUTRAL', 'Description'] = 'Internal transfer'

            df.loc[df['Direction'] == 'IN', 'Amount_currency'] = df['Source amount (after fees)']
            df.loc[df['Direction'] == 'OUT', 'Amount_currency'] = df['Target amount (after fees)'] * -1
            df.loc[df['Direction'] == 'NEUTRAL', 'Amount_currency'] = 0

            df.loc[df['Direction'] == 'IN', 'Currency'] = df['Source currency']
            df.loc[df['Direction'] == 'OUT', 'Currency'] = df['Target currency']
            df.loc[df['Direction'] == 'NEUTRAL', 'Currency'] = df['Target currency']

            # Create unique identifier
            df['UniqueID'] = df['ID']

            # Create bank identifier
            df['Bank'] = 'Wise'

            #create DKK_rate
            df['Amount'] = df.apply(calculate_target_currency_rate, axis=1, c=c)
            print(df['Amount'])
            df['Currency_Rate'] = df.apply(get_currency_rate, axis=1, c=c)

            selected_columns = ['Date', 'Description', 'Amount', 'Amount_currency', 'Currency', 'Currency_Rate', 'Bank', 'UniqueID']
            df_selected = df[selected_columns]

        elif bank == 'Lunar':
            delimiter = ','  # Change this delimiter as per the CSV format
            encoding = 'utf-8'  # Change encoding if necessary

            df = pd.read_csv(file_path, delimiter=delimiter, encoding=encoding)

            df = df.rename(columns={'Dato': 'Date', 'Tekst': 'Description', 'Beløb': 'Amount', 'Valuta': 'Currency'})
            df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d').dt.strftime('%d-%m-%Y')

            # Remove thousand separator (dot) and replace comma with dot for decimal delimiter in 'Amount' column for Danske Bank
            df['Amount'] = df['Amount'].str.replace('.', '').str.replace(',', '.', regex=False).astype(float)

            # Create unique identifier
            df['UniqueID'] = df.apply(lambda row: hashlib.md5(f"{row['Date']}_{row['Description']}_{row['Amount']}_{row['Saldo']}".encode()).hexdigest(), axis=1)

            # Create bank identifier
            df['Bank'] = 'Lunar'
            df['Amount_currency'] = df['Amount']
            df['Currency_Rate'] = 1

            selected_columns = ['Date', 'Description', 'Amount', 'Amount_currency', 'Currency', 'Currency_Rate', 'Bank', 'UniqueID']
            df_selected = df[selected_columns]

        elif bank == 'Skrill':
            delimiter = ','
            encoding = 'utf-8'

            df = pd.read_csv(file_path, delimiter=delimiter, encoding=encoding)

            df = df.rename(columns={'ID': 'UniqueID', 'Time (CET)': 'Date', 'Transaction Details': 'Description', 'Transaction Currency': 'Currency'})
            df['Date'] = pd.to_datetime(df['Date'], format='%d %b %y %H:%M').dt.strftime('%d-%m-%Y')

            # Convert columns to numeric and replace non-numeric values with NaN
            df['[-]'] = pd.to_numeric(df['[-]'], errors='coerce')
            df['[+]'] = pd.to_numeric(df['[+]'], errors='coerce')

            df['Amount'] = 0 - df['[-]'].fillna(0) + df['[+]'].fillna(0)

            df['Bank'] = 'Skrill'
            df['Amount_currency'] = df['Amount']
            df['Currency_Rate'] = 1

            selected_columns = ['Date', 'Description', 'Amount', 'Amount_currency', 'Currency', 'Currency_Rate', 'Bank', 'UniqueID']
            df_selected = df[selected_columns]


        else:
            raise ValueError("Unsupported bank")

        # Convert the DataFrame to a list of dictionaries
        print("Data extraction complete.")
        transactions = df_selected.to_dict(orient='records')

    except FileNotFoundError:
        print(f"File not found: {file_path}")
        transactions = []
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        transactions = []

    return transactions

def extract_pdf_data(file_path, bank):
    try:          
        if bank == 'Forbrugsforeningen':
            pdf_file = file_path
            dfs = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True, encoding='utf-8')
            
            # Initialize an empty DataFrame
            extracted_df = pd.DataFrame()
            
            # Iterate through the list of DataFrames and concatenate them
            for df in dfs:
                if 'Dato' in df.columns:
                    # Concatenate the DataFrame to the existing extracted_df
                    extracted_df = pd.concat([extracted_df, df], ignore_index=True)
            
            # Manipulate on pd.df
            extracted_df = extracted_df.rename(columns={'Dato': 'Date', 'Posteringstekst': 'Description', 'Beløb': 'Amount', 'Valuta': 'Currency'})
            extracted_df['Date'] = pd.to_datetime(extracted_df['Date'], format='%d/%m/%Y').dt.strftime('%d-%m-%Y')
            extracted_df['Amount'] = extracted_df['Amount'].str.replace('.', '').str.replace(',', '.', regex=False).astype(float)
            print(extracted_df['Amount'])
            extracted_df['UniqueID'] = extracted_df.apply(lambda row: hashlib.md5(f"{row['Date']}_{row['Description']}_{row['Amount']}".encode()).hexdigest(), axis=1)
            # Create bank identifier
            extracted_df['Bank'] = 'Forbrugsforeningen'
            extracted_df['Amount_currency'] = extracted_df['Amount']
            extracted_df['Currency_Rate'] = 1
            
            selected_columns = ['Date', 'Description', 'Amount', 'Amount_currency', 'Currency', 'Currency_Rate', 'Bank', 'UniqueID']
            df_selected = extracted_df[selected_columns]

        elif bank == 'Other bank':
            pass
        else:
            raise ValueError("Unsupported bank")

        # Convert the DataFrame to a list of dictionaries
        transactions = df_selected.to_dict(orient='records')
        return transactions

    except FileNotFoundError:
        print(f"File not found: {file_path}")
        transactions = []
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        transactions = []

    return 

def extract_xlsx_data(file_path, bank):
    try:
        if bank == 'Norwegian':
            # Specific settings for Bank Norwegian
            # Modify this section as per Bank Norwegian's format
            
            # Read the XLSX file into a pandas DataFrame with specific settings
            df = pd.read_excel(file_path)
            
            # Rename columns for consistency and readability
            df = df.rename(columns={'TransactionDate': 'Date', 'Text': 'Description',
                                    'Amount': 'Amount', 'Type': 'Type', 'Currency Amount': 'Amount_currency', 
                                    'Currency Rate': 'Currency_Rate'})

            # Data manipulation and formatting as needed for Bank Norwegian's XLSX format
            # You can add specific data manipulation steps here

            # Format the date column to 'DD-MM-YYYY' format
            df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y').dt.strftime('%d-%m-%Y')

            # Create unique identifier
            df['UniqueID'] = df.apply(lambda row: hashlib.md5(f"{row['Date']}_{row['Description']}_{row['Amount']}".encode()).hexdigest(), axis=1)

            # Create bank identifier
            df['Bank'] = 'Norwegian'

            selected_columns = ['Date', 'Description', 'Amount', 'Type', 'Amount_currency', 'Currency', 'Currency_Rate', 'UniqueID', 'Bank']
            df_selected = df[selected_columns]
            df_selected = df_selected[df_selected['Type'] != "Reserveret"]
            df_selected = df_selected[df_selected['Type'] != "Indbetaling"]
            df_selected = df_selected.drop('Type', axis=1)

        elif bank == 'Other Bank':
            # Specific settings for another bank (add more banks as needed)
            # Modify this section as per the other bank's format
            pass
        else:
            raise ValueError("Unsupported bank")

        # Convert the DataFrame to a list of dictionaries
        transactions = df_selected.to_dict(orient='records')

    except FileNotFoundError:
        print(f"File not found: {file_path}")
        transactions = []
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        transactions = []

    return transactions


### DATA TRANSFORMATION FUNCTIONS

def process_and_export_data(transactions, output_file_path):
    try:
        if os.path.isfile(output_file_path):
            # Load the existing CSV file
            existing_data = pd.read_csv(output_file_path, encoding='utf-8', sep=';', decimal=',')

            # Create a DataFrame from the new transactions
            new_data = pd.DataFrame(transactions)

            # Convert 'UniqueID' columns to strings in both DataFrames
            new_data['UniqueID'] = new_data['UniqueID'].astype(str)
            existing_data['UniqueID'] = existing_data['UniqueID'].astype(str)

            # Check for duplicates based on the 'UniqueID' column
            duplicates = new_data[new_data['UniqueID'].isin(existing_data['UniqueID'])]

            if not duplicates.empty:
                # Remove duplicates from the new_data DataFrame
                new_data = new_data[~new_data['UniqueID'].isin(existing_data['UniqueID'])]

                print(f"Removed {len(duplicates)} duplicate rows from new data.")

            # Concatenate the existing data and new data
            combined_data = pd.concat([existing_data, new_data], ignore_index=True)

            # Write the combined data to the same CSV file
            combined_data.to_csv(output_file_path, index=False, sep=";", decimal=",")

            print(f"Data appended to {output_file_path}")
        else:
            print(f"Creating a new CSV file: {output_file_path}")
            # If the file doesn't exist, create a new one
            df = pd.DataFrame(transactions)
            df.to_csv(output_file_path, index=False, encoding='utf-8', sep=";", decimal=",")

            print(f"Data exported to a new file: {output_file_path}")

    except Exception as e:
        print(f"An error occurred in process_and_export_data: {str(e)}")

### CATEGORIZATION FUNCTIONS

# Function to search for a category using the new description in categorization_rules.json
def search_category_by_description(new_description, categorization_rules):
    for rule in categorization_rules:
        if rule['keyword'].lower() in new_description.lower():
            return rule['category']
    return None

# Function to load categorization rules from a JSON file
def load_categorization_rules(file_path):
    try:
        with open(file_path, 'r') as json_file:
            categorization_rules = json.load(json_file)
        return categorization_rules
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return []
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return []

def append_categorization_rules(file_path, new_rules):
    existing_rules = load_categorization_rules(file_path)

    if existing_rules is None:
        existing_rules = []

    existing_rules.extend(new_rules)

    try:
        with open(file_path, 'w') as json_file:
            json.dump(existing_rules, json_file, indent=4)
        print(f"Categorization rules appended to {file_path}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def categorize_transactions(transactions, rules):
    categorized_transactions = []

    for transaction in transactions:
        categorized_transaction = transaction.copy()  # Create a copy of the original transaction

        # Iterate through the rules to find a match
        for rule in rules:
            keyword = rule['Keyword']
            main_category = rule['Main Category']
            sub_category = rule['Sub Category']

            # Check if the description starts with the keyword
            if transaction['Description'].lower().startswith(keyword.lower()):
                categorized_transaction['Main Category'] = main_category
                categorized_transaction['Sub Category'] = sub_category
                break  # Break the loop if a match is found

        # If no match is found, assign default categories
        if 'Main Category' not in categorized_transaction:
            categorized_transaction['Main Category'] = 'Uncategorized'
        if 'Sub Category' not in categorized_transaction:
            categorized_transaction['Sub Category'] = 'Uncategorized'

        # Append the categorized transaction to the list
        categorized_transactions.append(categorized_transaction)

    return categorized_transactions

def update_existing_category(output_csv_path, categorization_rules_path):
    try:
        # Load categorization rules from the JSON file
        with open(categorization_rules_path, 'r') as json_file:
            categorization_rules = json.load(json_file)

        # Load existing data from the output CSV file
        existing_data = pd.read_csv(output_csv_path, encoding='utf-8', sep=';', decimal=',')

        # Create a dictionary of categorization rules for faster lookup
        categorization_dict = {rule['Keyword'].lower(): rule for rule in categorization_rules}

        # Define a function to categorize descriptions using the same logic as the first function
        def categorize_description(desc):
            for keyword, rule in categorization_dict.items():
                if desc.lower().startswith(keyword.lower()):
                    return rule['Main Category'], rule['Sub Category']
            return 'Uncategorized', 'Uncategorized'

        # Apply the categorize_description function to the 'Description' column
        existing_data[['Main Category', 'Sub Category']] = existing_data['Description'].apply(categorize_description).apply(pd.Series)

        # Save the updated data back to the output CSV file
        existing_data.to_csv(output_csv_path, index=False, sep=";", decimal=",", encoding='utf-8')

        print(f"Categories updated in {output_csv_path}")

    except Exception as e:
        print(f"An error occurred in update_existing_category: {str(e)}")


### TESTING FUNCTIONS

def check_categories_and_duplicates(json_file_path):
    print(f"Checking file: {json_file_path}")

    try:
        # Load the JSON data from the file
        with open(json_file_path, 'r') as file:
            data = json.load(file)

        keywords_to_main_categories = {}
        keywords_to_sub_categories = {}
        conflicting_categories = {}

        for item in data:
            keyword = item.get('Keyword')
            main_category = item.get('Main Category')
            sub_category = item.get('Sub Category')

            if keyword is None:
                continue

            if keyword in keywords_to_main_categories:
                if keywords_to_main_categories[keyword] != main_category:
                    # This keyword has conflicting main categories
                    if keyword in conflicting_categories:
                        conflicting_categories[keyword].append(main_category)
                    else:
                        conflicting_categories[keyword] = [keywords_to_main_categories[keyword], main_category]
                else:
                    # Main category matches, check sub-categories
                    if keyword in keywords_to_sub_categories:
                        if keywords_to_sub_categories[keyword] != sub_category:
                            # This keyword has conflicting sub-categories
                            if keyword in conflicting_categories:
                                conflicting_categories[keyword].append(sub_category)
                            else:
                                conflicting_categories[keyword] = [keywords_to_sub_categories[keyword], sub_category]
                    else:
                        # Store sub-category for this keyword
                        keywords_to_sub_categories[keyword] = sub_category

            else:
                # Store main category for this keyword
                keywords_to_main_categories[keyword] = main_category

        if conflicting_categories:
            return conflicting_categories
        else:
            return {}

    except Exception as e:
        print(f"An error occurred in check_categories_and_duplicates: {str(e)}")
        return {}


def find_redundant_keywords(json_file_path):
    print(f"Checking file for redundant keywords: {json_file_path}")

    try:
        # Load the JSON data from the file
        with open(json_file_path, 'r') as file:
            data = json.load(file)

        keywords_count = {}
        redundant_keywords = {}

        for item in data:
            keyword = item.get('Keyword')

            if keyword is None:
                continue

            if keyword in keywords_count:
                # This keyword is redundant
                if keyword in redundant_keywords:
                    redundant_keywords[keyword].append(item)
                else:
                    redundant_keywords[keyword] = [keywords_count[keyword], item]
                keywords_count[keyword].append(item)  # Store all occurrences of the same keyword
            else:
                keywords_count[keyword] = [item]

        if redundant_keywords:
            return redundant_keywords
        else:
            return {}

    except Exception as e:
        print(f"An error occurred in find_redundant_keywords: {str(e)}")
        return {}

def main():
    # Specify the output file path for the CSV file
    output_file_path = 'C:/Users/smrie/OneDrive/60-69 Finance/62 Personal Book/AutoRegner/GIT/output.csv'

    #Create backup
    backup_csv_file(output_file_path)

    #Check if the CSV file exists before proceeding

    while is_file_open(output_file_path):
        print("please close the file")
        input("Press any key to continue")

    try:
        # Specify the path to the statement file (CSV, PDF, or XML)
        file_path = 'C:/Users/smrie/Downloads/eksport (1).pdf'  # Update with your file path
        bank = 'Forbrugsforeningen'

        # Specify the path to the categorization rules JSON file
        categorization_rules_path = 'categorization_rules.json'  # Replace with your actual path

        # Extract data from the statement file based on its format
        if file_path.endswith('.csv'):
            transactions = extract_csv_data(file_path, bank)
        elif file_path.endswith('.pdf'):
            transactions = extract_pdf_data(file_path, bank)
        elif file_path.endswith('.xlsx'):
            transactions = extract_xlsx_data(file_path, bank)
        else:
            raise ValueError("Unsupported file format")

        # Load categorization rules from the JSON file
        categorization_rules_list = load_categorization_rules(categorization_rules_path)

        # Categorize transactions
        pre_auto_categorize_transactions = categorize_transactions(transactions, categorization_rules_list)

        # Update existing
        update_existing_category(output_file_path, categorization_rules_path)

        # Perform further data processing or export the data to Excel
        process_and_export_data(pre_auto_categorize_transactions, output_file_path)

        #Post-processing of categories
        # Test the check_categories_and_duplicates function
        conflicting = check_categories_and_duplicates('categorization_rules.json')
        print("Conflicting Categories:", conflicting)
        redundant = find_redundant_keywords('categorization_rules.json')
        print("Redundant keywords: ", redundant)

    except Exception as e:
        print(f"An error occurred in the main function: {str(e)}")    


if __name__ == "__main__":

    #Continue with the main function
    main()
