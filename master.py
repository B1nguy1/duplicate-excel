import os
import pandas as pd


def write_to_excel(filepath, df):
    filepath = filepath + ".xlsx"
    if not os.path.isfile(filepath):
        writer = pd.ExcelWriter(filepath, engine='openpyxl')
        wb = writer.book
        df.to_excel(writer, index=False)
        wb.save(filepath)
    else:
        print(f"The file {filepath} already exists! Please create a new one.")


def delete_and_save_column(file_path, column_to_delete):
    file_path = file_path+".xlsx"
   
    df = pd.read_excel(file_path)

    if column_to_delete in df.columns:

        df.drop(columns=[column_to_delete], inplace=True)

        df.to_excel(file_path, index=False)
        print(f"Column '{column_to_delete}' deleted, and the data is saved to '{file_path}'.")
    else:
        print(f"Column '{column_to_delete}' not found in the DataFrame.")


def format_org_names(series):
    def process_string(s):
        if ' ' in s:
            return s.replace(" ", "").lower()
        else:
            return s.lower()
    return series.apply(process_string)


def find_duplicate_excel_values(old_excel_filepath, new_excel_filepath, rows, column):
    df = pd.read_excel(old_excel_filepath+'.xlsx', skiprows=rows, index_col=None)
    df["modified_org"] = df[column]
    org_count = df["modified_org"].value_counts()
    df["count"] = df["modified_org"].map(org_count)
   
    df_filtered = df[df["count"] > 1]
    
    if len(df_filtered.index > 0):
        df_filtered = df_filtered.sort_values(by=["modified_org"], ascending=False)
   
        df_filtered = df_filtered.loc[:, ~df_filtered.columns.str.contains('^Unnamed')]

        write_to_excel(new_excel_filepath, df_filtered)

        delete_and_save_column(new_excel_filepath, "modified_org")
    else:
        print("No occurences in the excel file!")