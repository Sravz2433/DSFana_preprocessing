import streamlit as st
import zipfile
import pandas as pd
from io import StringIO, BytesIO
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl import Workbook

st.title("📊  DSFANA Preprocessing - TXT Merger and Averager from ZIP")

uploaded_zip = st.file_uploader("Upload ZIP file containing .txt files", type=["zip"])

if uploaded_zip:
    st.success("✅ ZIP file uploaded!")

    # Read ZIP in memory
    zip_bytes = BytesIO(uploaded_zip.read())
    dataframes = {}
    row_counts = []

    with zipfile.ZipFile(zip_bytes, 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            if file_name.endswith('.txt'):
                with zip_ref.open(file_name) as file:
                    data = file.readlines()
                    modified_data = []
                    for line in data:
                        line_str = line.decode('utf-8')
                        modified_data.append(line_str.replace(' ', ',').strip())
                    modified_str = '\n'.join(modified_data)
                    df = pd.read_csv(StringIO(modified_str), header=None)
                    df = df.dropna(axis=1, how='all')
                    base_name = file_name.split('/')[-1].replace('.txt', '')
                    df[f"avg({base_name})"] = df.mean(axis=1)
                    row_counts.append(len(df))
                    dataframes[base_name] = df

    if len(set(row_counts)) != 1:
        st.error("❌ Not all files have the same number of rows.")
    else:
        # Create Excel in memory
        output = BytesIO()
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Combined Data"
        ws2 = wb.create_sheet("Averages")

        grey_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        center_align = Alignment(horizontal="center", vertical="center")
        red_font = Font(color="FF0000", bold=True)

        start_col = 2
        ws1.cell(row=1, column=1, value="SL No.")
        for df_name, df in dataframes.items():
            col_count = df.shape[1]
            col_start = start_col
            col_end = col_start + col_count - 1
            ws1.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
            ws1.cell(row=1, column=col_start).value = df_name
            ws1.cell(row=1, column=col_start).fill = grey_fill
            ws1.cell(row=1, column=col_start).alignment = center_align
            start_col += col_count

        # Write serial numbers
        for i in range(row_counts[0]):
            ws1.cell(row=i + 3, column=1, value=i + 1)

        # Write data
        start_col = 2
        for df in dataframes.values():
            for i, row in df.iterrows():
                for j, val in enumerate(row):
                    cell = ws1.cell(row=i + 3, column=start_col + j, value=val)
                    if j == len(row) - 1:
                        cell.font = red_font
            start_col += df.shape[1]

        # Sheet 2 - Averages
        ws2.append(["SL No."] + list(dataframes.keys()))
        for i in range(row_counts[0]):
            row = [i + 1]
            for df in dataframes.values():
                row.append(df.iloc[i, -1])
            ws2.append(row)
        for col in range(2, len(dataframes) + 2):
            for row in range(2, row_counts[0] + 2):
                ws2.cell(row=row, column=col).font = red_font

        wb.save(output)
        output.seek(0)

        st.success("✅ Successfully processed and generated Excel file.")
        st.download_button("📥 Download Combined Excel", output, file_name="combined_data.xlsx")
