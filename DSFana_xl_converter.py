import streamlit as st
import zipfile
import pandas as pd
from io import BytesIO, StringIO
import os
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl import Workbook

st.title("📊 DNA Structure Data Processor")
st.markdown("Process ZIP files of DNA structure data and generate formatted Excel reports")

# File upload section
uploaded_zip = st.file_uploader("Upload ZIP file with data files", type="zip")
uploaded_seq = st.file_uploader("Upload sequence.txt file", type="txt")

if uploaded_zip and uploaded_seq:
    if st.button("Process Files"):
        try:
            with st.spinner("Processing files..."):
                # Read sequence data
                sequence_text = uploaded_seq.read().decode("utf-8")
                sequence_lines = sequence_text.splitlines()
                sequence_ids = []
                sequences = []
                for i in range(0, len(sequence_lines), 2):
                    seq_id = sequence_lines[i].lstrip('>')
                    seq = sequence_lines[i+1] if i+1 < len(sequence_lines) else ''
                    sequence_ids.append(seq_id)
                    sequences.append(seq)

                # Process ZIP file
                dataframes = {}
                row_counts = []
                with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
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
                                base_name = os.path.splitext(os.path.basename(file_name))[0]
                                df[f"avg({base_name})"] = df.mean(axis=1)
                                row_counts.append(len(df))
                                dataframes[base_name] = df.round(5)

                # Validate row counts
                if len(set(row_counts)) != 1:
                    st.error("❌ Not all files have the same number of rows. Please check!")
                    st.stop()

                # Create Excel workbook in memory
                output = BytesIO()
                wb = Workbook()
                
                # [Rest of your existing Excel creation code here...]
                # Step 3: Create Excel Workbook
                wb = Workbook()
                ws1 = wb.active
                ws1.title = "Combined Data"
                ws2 = wb.create_sheet("Averages")

                # Define styles
                grey_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
                blue_fill = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type="solid")
                pink_fill = PatternFill(start_color="FCD5B4", end_color="FCD5B4", fill_type="solid")
                green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
                center_align = Alignment(horizontal="center", vertical="center")
                blue_font = Font(color="0813f8", bold=False)
                thin = Side(border_style="thin", color="000000")
                thick = Side(border_style="thick", color="000000")
                thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)
                thick_border = Border(left=thick, right=thick, top=thick, bottom=thick)
                thick_ud_border = Border(left=thin, right = thin , top = thick , bottom = thick)
                thick_lr_border = Border(left=thick, right=thick, top=thin, bottom=thin)
                
                # Step 4: Write to Sheet 1 (Merged headers and blue mean columns)
                # Step 4: Write to Sheet 1 (Merged headers and blue mean columns)
                start_col = 3  # First 2 columns are for sequence ID and sequence

                # Write "PARAMETERS" header (merged cells A1-B1)
                parameters_cell = ws1.cell(row=1, column=1, value="PARAMETERS")
                parameters_cell.fill = grey_fill
                parameters_cell.alignment = center_align
                parameters_cell.font = Font(bold=True)  
                parameters_cell.border = thick_border
                ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)

                for df_name, df in dataframes.items():
                    col_count = df.shape[1]
                    col_start = start_col
                    col_end = start_col + col_count - 1
                    ws1.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
                    ws1.cell(row=1, column=col_start).value = df_name
                    ws1.cell(row=1, column=col_start).fill = grey_fill
                    ws1.cell(row=1, column=col_start).alignment = center_align
                    ws1.cell(row=1, column=col_start).border = thin_border
                    start_col += col_count
                    

                max_rows = max([df.shape[0] for df in dataframes.values()])

                # Write sequence IDs and sequences to the worksheet starting from row 3
                for i in range(max_rows):
                    seq_id = sequence_ids[i] if i < len(sequence_ids) else ''
                    seq_value = sequences[i] if i < len(sequences) else ''
                    ids=ws1.cell(row=i + 3, column=1, value=seq_id)   
                    ids.fill = pink_fill
                    ids.border = thin_border
                    idv = ws1.cell(row=i + 3, column=2, value=seq_value)    # Sequence Value in column 2
                    idv.fill = blue_fill
                    idv.border = thin_border
                # Write headings in row 2
                seq_id = ws1.cell(row=2, column=1, value="Sequence ID")
                seq_id.font = Font(bold=True)
                seq_id.alignment = center_align
                seq_id.border = thin_border
                seq = ws1.cell(row=2, column=2, value="Sequence")
                seq.font = Font(bold=True)
                seq.alignment = center_align
                seq.border = thin_border
                ws1.column_dimensions['A'].width = 15  # Set column A to width 25
                ws1.column_dimensions['B'].width = 15  # Set column B to width 30



                col = 3  # Start from column 3 for dataframes
                for df in dataframes.values():
                    n_cols = df.shape[1]
                    # If your dataframe columns are numbered 0,1,2,...,Mean:
                    for i in range(n_cols - 1):  # All except last column (mean)
                        valu= ws1.cell(row=2, column=col, value=i)
                        valu.fill = green_fill
                        valu.border = thin_border
                        
                        col += 1
                    # Last column is mean
                    mean_cell = ws1.cell(row=2, column=col, value="Mean")
                    mean_cell.fill = green_fill
                    mean_cell.font = Font(bold=True)
                    mean_cell.alignment = center_align
                    mean_cell.border = thin_border

                    col += 1


                start_col = 3
                for df in dataframes.values():
                    for i, row in df.iterrows():
                        for j, val in enumerate(row):
                            cell = ws1.cell(row=i + 3, column=start_col + j, value=val)
                            if j == len(row) - 1:  # Last column (mean)
                                cell.font = blue_font
                                cell.border = thin_border
                    start_col += df.shape[1]
                             

                # Write headings in row 1
                ws2.cell(row=1, column=1, value="Sequence ID").font = Font(bold=True)
                ws2.cell(row=1, column=1).alignment = center_align
                ws2.cell(row=1, column=1).border = thin_border

                ws2.cell(row=1, column=2, value="Sequence").font = Font(bold=True)
                ws2.cell(row=1, column=2).alignment = center_align
                ws2.cell(row=1, column=2).border = thin_border

                # Optionally, write dataframe names as headings starting from column 3
                for col_idx, (name, df) in enumerate(dataframes.items(), start=3):
                    cell = ws2.cell(row=1, column=col_idx, value=f"{name}")
                    cell.font = Font(bold=True)
                    cell.alignment = center_align
                    cell.border = thin_border

                # Write sequence IDs and sequences to the worksheet starting from row 2
                max_rows = max([df.shape[0] for df in dataframes.values()])
                for i in range(max_rows):
                    cell_id = ws2.cell(row=i+2, column=1, value=sequence_ids[i] if i < len(sequence_ids) else '')
                    cell_id.fill = pink_fill
                    cell_id.border = thin_border

                    cell_seq = ws2.cell(row=i+2, column=2, value=sequences[i] if i < len(sequences) else '')
                    cell_seq.fill = blue_fill
                    cell_seq.border = thin_border

                # Write mean values from dataframes (starting from column 3)
                for col_idx, (name, df) in enumerate(dataframes.items(), start=3):
                    for row_idx in range(df.shape[0]):
                        cell = ws2.cell(row=row_idx+2, column=col_idx, value=df.iloc[row_idx, -1])
                        cell.border = thin_border

                # Ensure all cells in the used range have borders (for empty cells)
                total_cols = 2 + len(dataframes)
                for row in range(1, max_rows + 2):  # +2 because headings in row 1
                    for col in range(1, total_cols + 1):
                        ws2.cell(row=row, column=col).border = thin_border

                # Adjust column widths for better readability
                for col in ws2.columns:
                    max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                    ws2.column_dimensions[col[0].column_letter].width = max_length + 2

                # ... (include all your existing Excel formatting code)

                # Save workbook to BytesIO object
                wb.save(output)
                output.seek(0)

                st.success("✅ Processing complete!")
                st.download_button(
                    label="Download Excel Report",
                    data=output,
                    file_name="dna_structure_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
