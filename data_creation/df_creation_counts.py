import pandas as pd

def create_counts_df(workbook):
    # Load the WCAG data into a Pandas dataframe
    df = pd.read_csv("wcag_data.csv")
    df = df.set_index("WCAG")

    print(df.head())

    # Create a dictionary to store the counts of each WCAG SC#
    wcag_sc_counts = {}

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        wcag_sc_column_index = None

        # Search for the column named "WCAG SC#"
        for column in sheet.iter_cols():
            for cell in column:
                if cell.value == 'WCAG SC#':
                    wcag_sc_column_index = cell.column_letter
                    break
            if wcag_sc_column_index is not None:
                break

        # If the column is found, iterate over all cells in that column
        if wcag_sc_column_index is not None:
            for cell in sheet[f'{wcag_sc_column_index}2:{wcag_sc_column_index}{sheet.max_row}']:
                wcag_sc = cell[0].value
                if wcag_sc is not None:
                    try:
                        df["No_of_occurence"][wcag_sc] = df["No_of_occurence"][wcag_sc] + 1
                    except:
                        print("Error Marked Wrongly Moving Forward")
                    wcag_sc_counts[wcag_sc] = wcag_sc_counts.get(wcag_sc, 0) + 1

    # Filter the dataframe to only include WCAG SCs that appear in the workbook
    new_df = df[['WCAG_SC', 'No_of_occurence']][df['No_of_occurence'] > 0].reset_index(drop=True)

    return new_df
