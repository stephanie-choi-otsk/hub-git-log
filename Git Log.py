import subprocess
import pandas as pd
import os
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment

def git_pull(repo_path):
    # Change directory to the local repository
    try:
        subprocess.check_output(["git", "rev-parse"], cwd=repo_path, stderr=subprocess.STDOUT, text=True)
    except subprocess.CalledProcessError:
        print(f"Error: The specified path '{repo_path}' is not a Git repository.")
        return False

    # Pull the latest changes from the remote repository
    subprocess.run(["git", "pull"], cwd=repo_path)
    return True

def get_merge_commit_details(repo_path, pull_request_branch):
    # Perform git pull to ensure you have the latest changes
    if not git_pull(repo_path):
        return None

    # Get the list of merge commit details on the pull request branch
    merge_commit_details = subprocess.check_output(["git", "log", "--merges", "--pretty=%H|%ct|%s", "--grep=^Merge pull request"], cwd=repo_path, text=True).splitlines()

    return merge_commit_details

def parse_merge_commit_details(details):
    # Parse merge commit details into a DataFrame
    parsed_details = []
    for line in details:
        commit_hash, timestamp, pr_info = line.split('|', 2)
        pr_number = pr_info.split('#', 1)[1].split(' ', 1)[0] if '#' in pr_info else None
        merge_date = pd.to_datetime(int(timestamp), unit='s')
        
        # Extract PR link from the commit message
        pr_link = find_pr_link(pr_info)
        
        parsed_details.append({'Merge Date': merge_date, 'PR Number': pr_number, 'PR Link': pr_link, 'Merge Commit Hash': commit_hash})
    
    return parsed_details

def find_pr_link(message):
    # Find GitHub pull request URL in the commit message using specific pattern
    pr_number_match = re.search(r'#(\d+)', message)
    if pr_number_match:
        pr_number = pr_number_match.group(1)
        return f'https://github.com/ocean-network-express/LOOKML_one_hub/pull/{pr_number}'

    return None

def export_to_excel(merge_commit_details, excel_path):
    # Create a DataFrame with merge commit details
    df = pd.DataFrame(merge_commit_details)

    # Reorder the columns
    df = df[['Merge Date', 'PR Number', 'PR Link', 'Merge Commit Hash']]

    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1', startrow=1, header=False)

        # Get the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Get the dimensions of the DataFrame
        num_rows, num_cols = df.shape

        # Create a list of column headers, to use in add_table()
        column_settings = [{'header': column} for column in df.columns]

        # Add the Excel table structure. Pandas will add the data.
        worksheet.add_table(0, 0, num_rows, num_cols - 1, {'columns': column_settings})

        # Get the max column width in characters
        max_len = max([len(str(column)) if column else 0 for column in df.columns] + [df.index.name or 0])
        
        # Adjust the column width
        for i, col in enumerate(df.columns):
            max_len = max(max_len, df[col].astype(str).map(len).max())
            worksheet.set_column(i, i, max_len)

        # Make the URL clickable (hyperlink)
        for i, url in enumerate(df['PR Link']):
            if url:
                worksheet.write_url(i + 1, 2, url, string='Link')

    print(f"Merge commit details exported to Excel: {excel_path}")

# Example usage
repo_path = r'C:\Users\stephani.choi\DDE\LOOKML_one_hub'  # Replace with the actual path to your local Git repository
pull_request_branch = 'pull_request_branch'  # Replace with the actual pull request branch name
excel_path = r'C:\Users\stephani.choi\Documents\merge_commit_details.xlsx'  # Replace with the desired Excel file path

merge_commit_details_raw = get_merge_commit_details(repo_path, pull_request_branch)

if merge_commit_details_raw:
    merge_commit_details_parsed = parse_merge_commit_details(merge_commit_details_raw)
    export_to_excel(merge_commit_details_parsed, excel_path)
