import pandas as pd
import os

def load_data(file_path, sheet_name):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    
    try:
        data = pd.read_excel(file_path, sheet_name=sheet_name)
        return data
    except Exception as e:
        raise Exception(f"Error loading Excel file: {e}")

def categorize_aging_buckets(data, dpd_column='dpd'):
    # Define the aging bucket categories based on DPD
    bins = [-1, 0, 30, 60, 90, float('inf')]
    labels = ['M0: 0 days', 'M1: 1-30 days', 'M2: 31-60 days', 'M3: 61-90 days', '>M3: >90 days']
    
    # Create a new 'Aging Bucket' column
    data['Aging Bucket'] = pd.cut(data[dpd_column], bins=bins, labels=labels)
    return data

def summarize_aging_buckets(data):

    # Group data by aging buckets and calculate summary
    summary = data.groupby('Aging Bucket', observed=False).agg(
        Number_of_Accounts=('agrt no.', 'count'),
        Total_GL_Balance=('gl bal', 'sum')
    ).reset_index()
    return summary

def main():
    # Define the file path and sheet name
    file_path = r'C:/Users/user/Desktop/hp_data_output.xlsx'
    sheet_name = 'HP Aging'
    
    # Load the data
    hp_aging_data = load_data(file_path, sheet_name)
    
    # Categorize the data into aging buckets
    hp_aging_data = categorize_aging_buckets(hp_aging_data)
    
    # Summarize the aging bucket data
    aging_summary = summarize_aging_buckets(hp_aging_data)
    
    # Display the summarized aging bucket data
    print(aging_summary)

if __name__ == "__main__":
    main()
