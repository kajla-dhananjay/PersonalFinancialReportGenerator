import gspread
import argparse
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import sys


def generate_sankey_data(df):
    grouped_data = df.groupby('H1')['Amount'].sum().reset_index()
    inflows_unsorted = grouped_data[grouped_data['Amount'] < 0]
    inflows = inflows_unsorted.sort_values(by=['Amount'], ascending=[True])

    total_income = abs(inflows['Amount'].sum()).round(2)

    grouped_outflows_h1_tmp = grouped_data[grouped_data['Amount'] >= 0]
    grouped_outflows_h1_unsorted = grouped_outflows_h1_tmp[~grouped_outflows_h1_tmp['H1'].isin(inflows['H1'])]
    grouped_outflows_h1 = grouped_outflows_h1_unsorted.sort_values(by=['Amount'], ascending=[False])

    h1_order = pd.CategoricalDtype(categories=grouped_outflows_h1['H1'], ordered=True)

    total_spent = grouped_outflows_h1['Amount'].sum()
    total_saved = total_income - total_spent

    grouped_outflows_h2_all = df.groupby(['H1', 'H2'])['Amount'].sum().reset_index()
    grouped_outflows_h2_tmp = grouped_outflows_h2_all.query('Amount > 0')
    grouped_outflows_h2_unsorted = grouped_outflows_h2_tmp[~grouped_outflows_h2_tmp['H1'].isin(inflows['H1'])]
    grouped_outflows_h2_sort1 = grouped_outflows_h2_unsorted.sort_values(by='H1', ascending=True)
    grouped_outflows_h2_sort1['H1'] = grouped_outflows_h2_sort1['H1'].astype(h1_order)
    grouped_outflows_h2 = grouped_outflows_h2_sort1.sort_values(by=['H1', 'Amount'], ascending=[True, False])

    grouped_outflows_h3_all = df.groupby(['H1', 'H3'])['Amount'].sum().reset_index()
    grouped_outflows_h3_tmp = grouped_outflows_h3_all.query('Amount > 0')
    grouped_outflows_h3_unsorted = grouped_outflows_h3_tmp[~grouped_outflows_h3_tmp['H1'].isin(inflows['H1'])]
    grouped_outflows_h3_sort1 = grouped_outflows_h3_unsorted.sort_values(by='H1', ascending=True)
    grouped_outflows_h3_sort1['H1'] = grouped_outflows_h3_sort1['H1'].astype(h1_order)
    grouped_outflows_h3 = grouped_outflows_h3_sort1.sort_values(by=['H1', 'Amount'], ascending=[True, False])

    links_str_inflows = '\n'.join(
        f'{inflows["H1"].iloc[i]} [{abs(inflows["Amount"].iloc[i]):.2f}] Income' for i in range(len(inflows)))
    links_str_outflows = '\n'.join([f'Income [{total_spent:.2f}] Expenses', f'Income [{total_saved:.2f}] Savings'])

    links_str_outflows_h1 = '\n'.join(
        f'Expenses [{grouped_outflows_h1["Amount"].iloc[i]:.2f}] {grouped_outflows_h1["H1"].iloc[i]} ' for i in
        range(len(grouped_outflows_h1)))
    links_str_outflows_h2 = '\n'.join(
        f'{grouped_outflows_h2["H1"].iloc[i]} [{grouped_outflows_h2["Amount"].iloc[i]:.2f}] {grouped_outflows_h2["H2"].iloc[i]}'
        for i in range(len(grouped_outflows_h2)))
    links_str_outflows_h3 = '\n'.join(
        f'{grouped_outflows_h3["H1"].iloc[i]} [{grouped_outflows_h3["Amount"].iloc[i]:.2f}] {grouped_outflows_h3["H3"].iloc[i]} '
        for i in range(len(grouped_outflows_h3)))

    sankeyList_2 = [links_str_inflows, links_str_outflows, links_str_outflows_h1,
                    links_str_outflows_h2]  # , links_str_outflows_h3]
    sankeyList_3 = [links_str_inflows, links_str_outflows, links_str_outflows_h1,
                    links_str_outflows_h3]  # , links_str_outflows_h3]
    sankeymatic_input_2 = '\n'.join(sankeyList_2)
    sankeymatic_input_3 = '\n'.join(sankeyList_3)

    return [sankeymatic_input_2, sankeymatic_input_3]


def download_google_sheet_range(sheet_url, sheet_name, start_row, end_row, start_col, end_col):
    # Set up credentials
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credential_file, scope)
    gc = gspread.authorize(credentials)

    # Open the Google Sheet
    workbook = gc.open_by_url(sheet_url)
    worksheet = workbook.worksheet(sheet_name)

    # Get the data in the specified range
    cell_range = worksheet.range(start_row, start_col, end_row, end_col)
    data = [cell.value for cell in cell_range]

    # Reshape the data into a 2D list
    num_rows = end_row - start_row + 1
    num_cols = end_col - start_col + 1
    data_2d = [data[i:i + num_cols] for i in range(0, len(data), num_cols)]

    # Convert the 2D list to a DataFrame
    df = pd.DataFrame([row for row in data_2d if any(row)])
    df.columns = ['Amount', 'H1', 'H2', 'H3']
    df['Amount'] = df['Amount'].str.replace('₹', '').str.replace(',', '').astype(float)
    return df


def usage():
    print("Usage: python3 script.py [-m MONTH] [-o OUTPUT_FILE]")
    sys.exit(1)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='GenerateReport')
    parser.add_argument('-o', type=str, default='out.txt', help='Output file name (default: out.txt)')
    parser.add_argument('-m', type=str, default='Jan', help='Report Generation Month')
    sheet_name = 'Jan'
    output_file = 'Jan'
    try:
        args = parser.parse_args()

        output_file = args.o
        sheet_name = args.m

    except argparse.ArgumentError as e:
        print(f"Error: {e}")
        usage()

    sheet_url = 'https://docs.google.com/spreadsheets/yourGoogleSheetUrl/edit'
    credential_file = '../../keys/keys.json'
    start_row = 10 # CFG: Your start row here
    end_row = 1024 # CFG: Your end row here
    start_col = 2 # CFG: Your start col here
    end_col = 5 # CFG: Your end col here
    dff = download_google_sheet_range(sheet_url, sheet_name, start_row, end_row, start_col, end_col)
    sankeymatic_input = generate_sankey_data(dff)
    print(sankeymatic_input[0])
    print('\n\n')
    print(sankeymatic_input[1])
    with open(output_file + '_1.txt', 'w') as file:
        print(sankeymatic_input[0], file=file)
    with open(output_file + '_2.txt', 'w') as file:
        print(sankeymatic_input[1], file=file)
