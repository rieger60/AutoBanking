# Import libraries
import pandas as pd
import json
import hashlib
import os
import shutil
import time

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

            selected_columns = ['Date', 'Description', 'Amount', 'UniqueID']
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
            df['Amount'] = ''
            df['Currency'] = ''

            # Set values based on 'Direction' column
            df.loc[df['Direction'] == 'IN', 'Description'] = df['Source name']
            df.loc[df['Direction'] == 'OUT', 'Description'] = df['Target name']
            df.loc[df['Direction'] == 'NEUTRAL', 'Description'] = 'Internal transfer'

            df.loc[df['Direction'] == 'IN', 'Amount'] = df['Source amount (after fees)']
            df.loc[df['Direction'] == 'OUT', 'Amount'] = df['Target amount (after fees)'] * -1
            df.loc[df['Direction'] == 'NEUTRAL', 'Amount'] = 0

            df.loc[df['Direction'] == 'IN', 'Currency'] = df['Source currency']
            df.loc[df['Direction'] == 'OUT', 'Currency'] = df['Target currency']
            df.loc[df['Direction'] == 'NEUTRAL', 'Currency'] = df['Target currency']

            # Create unique identifier
            df['UniqueID'] = df['ID']

            # Create bank identifier
            df['Bank'] = 'Wise'

            selected_columns = ['Date', 'Description', 'Amount', 'Currency', 'UniqueID']
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

# Function to extract data from a PDF file using Tabula
def extract_pdf_data(file_path, bank):
    # Implement this function to extract data from a PDF file using Tabula
    pass

import pandas as pd

def extract_xlsx_data(file_path, bank):
    try:
        if bank == 'Norwegian':
            # Specific settings for Bank Norwegian
            # Modify this section as per Bank Norwegian's format
            
            # Read the XLSX file into a pandas DataFrame with specific settings
            df = pd.read_excel(file_path, usecols=['TransactionDate', 'Text', 'Type', 'Amount'])
            
            # Rename columns for consistency and readability
            df = df.rename(columns={'TransactionDate': 'Date', 'Text': 'Description',
                                    'Amount': 'Amount', 'Type': 'Type'})

            # Data manipulation and formatting as needed for Bank Norwegian's XLSX format
            # You can add specific data manipulation steps here

            # Format the date column to 'DD-MM-YYYY' format
            df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y').dt.strftime('%d-%m-%Y')

            # Create unique identifier
            df['UniqueID'] = df.apply(lambda row: hashlib.md5(f"{row['Date']}_{row['Description']}_{row['Amount']}".encode()).hexdigest(), axis=1)

            # Create bank identifier
            df['Bank'] = 'Norwegian'

            selected_columns = ['Date', 'Description', 'Amount', 'Type', 'UniqueID', 'Bank']
            df_selected = df[selected_columns]
            df_selected = df_selected[df_selected['Type'] != "Reserveret"]
            df_selected = df_selected[df_selected['Type'] != "Indbetaling"]

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


# Function to categorize transactions based on rules
def categorize_transactions(transactions, rules):
    categorized_transactions = []

    for transaction in transactions:
        categorized_transaction = transaction.copy()  # Create a copy of the original transaction

        # Iterate through the rules to find a match
        for rule in rules:
            keyword = rule['keyword']
            category = rule['category']

            # Check if the keyword is in the transaction description
            if keyword.lower() in transaction['Description'].lower():
                categorized_transaction['Category'] = category
                break  # Break the loop if a match is found

        # If no match is found, assign a default category
        if 'Category' not in categorized_transaction:
            categorized_transaction['Category'] = 'Uncategorized'

        # Append the categorized transaction to the list
        categorized_transactions.append(categorized_transaction)

    return categorized_transactions

import os
import pandas as pd

def process_and_export_data(transactions, output_file_path):
    try:
        if os.path.isfile(output_file_path):
            # Load the existing CSV file
            existing_data = pd.read_csv(output_file_path)

            # Create a DataFrame from the new transactions
            new_data = pd.DataFrame(transactions)

            # Check for duplicates based on the 'UniqueID' column
            duplicates = new_data[new_data['UniqueID'].isin(existing_data['UniqueID'])]

            if not duplicates.empty:
                # Remove duplicates from the new_data DataFrame
                new_data = new_data[~new_data['UniqueID'].isin(existing_data['UniqueID'])]

                print(f"Removed {len(duplicates)} duplicate rows from new data.")

            # Concatenate the existing data and new data
            combined_data = pd.concat([existing_data, new_data], ignore_index=True)

            # Write the combined data to the same CSV file
            combined_data.to_csv(output_file_path, index=False)

            print(f"Data appended to {output_file_path}")
        else:
            print(f"Creating a new CSV file: {output_file_path}")
            # If the file doesn't exist, create a new one
            df = pd.DataFrame(transactions)
            df.to_csv(output_file_path, index=False, encoding='utf-8')

            print(f"Data exported to a new file: {output_file_path}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")



# Function to search for a category using the new description in categorization_rules.json
def search_category_by_description(new_description, categorization_rules):
    for rule in categorization_rules:
        if rule['keyword'].lower() in new_description.lower():
            return rule['category']
    return None

# Function to find and handle uncategorized transactions:
# - Identify transactions that were not categorized.
# - Update the original description to match a categorized text if possible.
# - Create a new category for uncategorized transactions if necessary.
# Function to find and handle uncategorized transactions:
# - Identify transactions that were not categorized.
# - Update the original description to match a categorized text if possible.
# - Create a new category for uncategorized transactions if necessary.
def find_uncategorized_transactions(categorized_transactions, categorization_rules):
    try:
        for transaction in categorized_transactions:
            if 'Category' not in transaction or transaction['Category'] == 'Uncategorized':
                print("Found an uncategorized transaction:")
                print(f"Transaction Description: {transaction['Description']}")
                print(f"Transaction Amount: {transaction['Amount']}")
                print(f"Transaction Date: {transaction['Date']}")

                # Prompt the user for input
                while True:
                    choice = input("Select an option:\n1. Add a new category\n2. Change description text\n3. Skip (leave uncategorized)\nChoice (1/2/3): ")

                    if choice == '1':
                        # Add a new category using user input
                        keyword = input("Enter a keyword: ")
                        category = input("Enter a category: ")

                        # Create a new rule dictionary
                        new_rule = {'keyword': keyword, 'category': category}

                        # Append the new rule to the existing categorization rules
                        append_categorization_rules('categorization_rules.json', [new_rule])

                        # Update the transaction's category
                        transaction['Category'] = category

                        print("Category added to the transaction.")
                        break
                    elif choice == '2':
                        # Update the transaction's description
                        new_description = input("Enter a new description: ")
                        transaction['Description'] = new_description

                        # Search for a category using the new description
                        found_category = search_category_by_description(new_description, categorization_rules)

                        if found_category:
                            # Update the transaction's category
                            transaction['Category'] = found_category
                            transaction['Description'] = new_description

                            # Print the updated transaction information
                            print("Updated Transaction Description:")
                            print(f"Transaction Description: {transaction['Description']}")
                            print(f"Transaction Amount: {transaction['Amount']}")
                            print(f"Transaction Date: {transaction['Date']}")
                        else:
                            print("No matching category found. Please add category: ")
                            category = input("Enter a category name: ")
                            new_rule = {'keyword': new_description, 'category': category}
                            transaction['Category'] = category

                            print("Category added to the transaction.")
                        break
                    elif choice == '3':
                        print("Transaction skipped (left uncategorized).")
                        break
                    else:
                        print("Invalid choice. Please select 1, 2, or 3.")

    except Exception as e:
        logger.error(f"An error occurred in find_uncategorized_transactions: {str(e)}")

    return categorized_transactions



    # Remove the processed transactions from the main list
    for transaction in transactions_to_remove:
        transactions.remove(transaction)

    return categorized_transactions


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
        file_path = 'C:/Users/smrie/Downloads/Statement (2).xlsx'  # Update with your file path
        bank = 'Norwegian'

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

        # Find and collect uncategorized transactions
        manual_categorize_transactions = find_uncategorized_transactions(pre_auto_categorize_transactions, categorization_rules_list)

        # Perform further data processing or export the data to Excel
        process_and_export_data(pre_auto_categorize_transactions, output_file_path)

    except Exception as e:
        print(f"An error occurred in the main function: {str(e)}")    


if __name__ == "__main__":

    #Continue with the main function
    main()
