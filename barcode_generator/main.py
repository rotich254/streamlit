import streamlit as st
import pandas as pd
from io import BytesIO
from barcode import Code128
from barcode.writer import ImageWriter
import xlsxwriter

def generate_barcode(value):
    barcode = Code128(str(value), writer=ImageWriter())
    buffer = BytesIO()
    barcode.write(buffer)
    return buffer.getvalue()

def main():
    st.title("Excel Barcode Generator")

    # File upload
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Load the Excel file with the first row as the header
        df = pd.read_excel(uploaded_file)

        # Display the DataFrame
        st.write("Data Preview:")
        st.dataframe(df)

        # Select a column to generate barcodes
        barcode_column = st.selectbox("Select a column to generate barcodes", df.columns)

        # Ask for the name of the new column for barcodes
        new_column_name = st.text_input("Enter the name for the new column where barcodes will be stored", "Barcodes")

        if st.button("Generate and Download"):
            # Create a copy of the dataframe
            df_copy = df.copy()

            # Add the new column to the DataFrame
            df_copy[new_column_name] = ""

            # Generate barcodes and store them in a dictionary
            barcodes = {}
            for index, value in df_copy[barcode_column].items():
                barcodes[index] = generate_barcode(value)

            # Save the modified DataFrame to a new Excel file
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_copy.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                # Insert barcode images into the Excel file and adjust cell sizes
                for index, image_data in barcodes.items():
                    image_stream = BytesIO(image_data)
                    worksheet.insert_image(index + 1, df_copy.columns.get_loc(new_column_name), '', {'image_data': image_stream, 'x_scale': 0.5, 'y_scale': 0.5})
                    worksheet.set_row(index + 1, 60)  # Adjust row height
                worksheet.set_column(df_copy.columns.get_loc(new_column_name), df_copy.columns.get_loc(new_column_name), 20)  # Adjust column width

            # Provide a download link for the new Excel file
            output.seek(0)
            st.download_button(
                label="Download Excel file with Barcodes",
                data=output,
                file_name="excel_with_barcodes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
