import pandas as pd
import sqlite3
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_excel_data(file_path, sheet_name=0):
    """
    Load data from an Excel file.
    """
    try:
        # Determine the correct engine based on file extension
        if file_path.endswith('.xls'):
            data = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')
        elif file_path.endswith('.xlsx'):
            data = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        else:
            logging.error(f"Unsupported file format: {file_path}")
            return None

        # Remove 'Unnamed' columns
        data = data.loc[:, ~data.columns.str.contains('^Unnamed')]

        logging.info(f"Data loaded successfully from {file_path}")
        return data
    except Exception as e:
        logging.error(f"Error loading data from {file_path}: {e}")
        return None

def standardize_dates(df, date_columns):
    """
    Standardize date columns to the format YYYY-MM-DD.
    """
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
    return df

def clean_hp_aging_data(df):
    """
    Clean and preprocess HP Aging data.
    """
    df.columns = df.columns.str.strip().str.lower()

    # Standardizing date columns
    date_columns = ['submission date', 'approval date']
    df = standardize_dates(df, date_columns)

    # Handle numeric and categorical columns
    numeric_columns = ['loan amt', 'mthly instal', 'arrears amt']
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Fill missing values for categorical columns
    categorical_columns = ['gender', 'dealer id', 'occupation']
    for col in categorical_columns:
        if col in df.columns:
            df[col] = df[col].fillna('Unknown')

    # Remove outliers in age
    if 'age' in df.columns:
        df = df[(df['age'] > 0) & (df['age'] <= 100)]

    # Remove duplicate rows
    df = df.drop_duplicates()

    logging.info("HP Aging data cleaned successfully")
    return df

def clean_hp_os_data(df):
    """
    Clean and preprocess HP OS data.
    """
    df.columns = df.columns.str.strip().str.lower()

    # Standardizing date columns
    date_columns = ['agrt date', 'last paid date']
    df = standardize_dates(df, date_columns)

    # Remove duplicate rows
    df = df.drop_duplicates()

    logging.info("HP OS data cleaned successfully")
    return df

def create_db_connection(db_name):
    """
    Create a SQLite database connection.
    """
    try:
        conn = sqlite3.connect(db_name)
        logging.info(f"Connected to the SQLite database: {db_name}")
        return conn
    except sqlite3.Error as e:
        logging.error(f"Failed to connect to database: {e}")
        return None

def insert_data_to_sql(conn, df, table_name):
    """
    Insert data into a SQLite table.
    """
    try:
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        logging.info(f"Data successfully inserted into the {table_name} table.")
    except Exception as e:
        logging.error(f"Error inserting data into {table_name} table: {e}")

def verify_data_insertion(conn, table_name):
    """
    Verify that data has been inserted into the SQLite table.
    """
    try:
        query = f"SELECT * FROM {table_name} LIMIT 5"
        df = pd.read_sql_query(query, conn)
        logging.info(f"Data in {table_name} (first 5 rows):\n{df}")
    except Exception as e:
        logging.error(f"Error verifying data in {table_name}: {e}")

def save_data_to_excel(hp_aging_data, hp_os_data, output_file):
    """
    Save the cleaned data to an Excel file with separate sheets for HP Aging and HP OS data.
    """
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            if hp_aging_data is not None:
                hp_aging_data.to_excel(writer, sheet_name='HP Aging', index=False)
            if hp_os_data is not None:
                hp_os_data.to_excel(writer, sheet_name='HP OS', index=False)
        logging.info(f"Data successfully saved to {output_file}")
    except Exception as e:
        logging.error(f"Error saving data to Excel file: {e}")

def main():
    # File paths
    file_aging = r'C:\Users\user\Desktop\geniemyseniorbianalystassessment\Data 3 (Hp Aging).xlsx'
    file_os = r'C:\Users\user\Desktop\geniemyseniorbianalystassessment\Data 4 (Hp OS).xlsx'
    output_file = r'C:\Users\user\Desktop\hp_data_output.xlsx'
    db_name = "hp_database.db"

    # Load and clean HP Aging data
    hp_aging_data = load_excel_data(file_aging)
    if hp_aging_data is not None:
        hp_aging_data = clean_hp_aging_data(hp_aging_data)

    # Load and clean HP OS data
    hp_os_data = load_excel_data(file_os)
    if hp_os_data is not None:
        hp_os_data = clean_hp_os_data(hp_os_data)

    # Create a SQLite database connection
    conn = create_db_connection(db_name)
    if conn:
        # Insert cleaned data into the SQLite database
        if hp_aging_data is not None:
            insert_data_to_sql(conn, hp_aging_data, "hp_aging")
        if hp_os_data is not None:
            insert_data_to_sql(conn, hp_os_data, "hp_os")

        # Verify the data insertion
        verify_data_insertion(conn, "hp_aging")
        verify_data_insertion(conn, "hp_os")

        # Close the database connection
        conn.close()

    # Save cleaned data to an Excel file
    save_data_to_excel(hp_aging_data, hp_os_data, output_file)

if __name__ == "__main__":
    main()
