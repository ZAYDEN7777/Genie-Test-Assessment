import pandas as pd

def load_and_preprocess_data(file_path):
    """
    Loads the dataset from the provided file path and preprocesses it by
    calculating the number of months completed for each loan.
    """
    # Load the Excel data
    df = pd.read_excel(file_path, sheet_name='Sheet')
    
    # Calculate how many months the loan has been active
    df['Months Completed'] = (pd.to_datetime(df['Last Paid Date']).dt.year - pd.to_datetime(df['Agrt Date']).dt.year) * 12 + \
                             (pd.to_datetime(df['Last Paid Date']).dt.month - pd.to_datetime(df['Agrt Date']).dt.month)
    
    # Calculate the percentage of installment progress
    df['Installment Progress'] = df['Months Completed'] / df['Tenor']
    
    # Filter out loans that have negative arrears (invalid or repaid in advance)
    df_active = df[df['Arrears'] >= 0]
    
    return df_active


def calculate_risk_categories(df):
    """
    Categorizes loans into risk categories based on the number of months overdue.
    Adds a 'Risk Category' column to the DataFrame.
    """
    # Define risk categories based on months overdue
    df['Risk Category'] = pd.cut(df['Mth Due'], bins=[-float('inf'), 0, 1, 2, float('inf')],
                                 labels=['Current', '1 Month Overdue', '2 Months Overdue', '3+ Months Overdue'])
    return df


def generate_summary_statistics(df):
    """
    Generates a summary of key metrics, including the total number of active loans, 
    total arrears, percentage of overdue loans, and installment progress summary.
    """
    # Summarize key metrics
    summary = {
        'Total Active Loans': len(df),
        'Total Arrears': df['Arrears'].sum(),
        'Average Arrears per Loan': df['Arrears'].mean(),
        'Percentage of Overdue Loans': (df['Arrears'] > 0).mean() * 100,
        'Average Installment Progress': df['Installment Progress'].mean() * 100  # Percentage
    }
    
    # Create installment progress categories
    installment_bins = pd.cut(df['Installment Progress'], bins=[0, 0.25, 0.5, 0.75, 1.0],
                              labels=['0-25%', '25-50%', '50-75%', '75-100%'])
    installment_summary = installment_bins.value_counts().sort_index()
    
    # Create a risk category summary
    risk_summary = df['Risk Category'].value_counts().sort_index()
    
    return summary, installment_summary, risk_summary


def display_results(summary, installment_summary, risk_summary):
    """
    Prints the results of the analysis in a readable format.
    """
    print("Company Risk Profile Summary:")
    for key, value in summary.items():
        print(f"{key}: {value}")
    
    print("\nInstallment Progress Summary:")
    print(installment_summary)
    
    print("\nRisk Category Summary:")
    print(risk_summary)


def main(file_path):
    """
    Main function to execute the workflow. Loads the data, calculates risk categories,
    generates summary statistics, and displays the results.
    """
    # Step 1: Load and preprocess data
    df_active = load_and_preprocess_data(file_path)
    
    # Step 2: Calculate risk categories
    df_active = calculate_risk_categories(df_active)
    
    # Step 3: Generate summary statistics
    summary, installment_summary, risk_summary = generate_summary_statistics(df_active)
    
    # Step 4: Display the results
    display_results(summary, installment_summary, risk_summary)


# Execute the main function with the appropriate file path
file_path = r'C:\Users\user\Desktop\geniemyseniorbianalystassessment\Data 4 (Hp OS).xlsx'  # Adjust the file path as necessary
main(file_path)
