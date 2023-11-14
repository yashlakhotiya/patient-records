import os

from utils import *


def main():
    pdf_directory = os.path.expanduser('~/Documents/patient-records/pdf')  # Expand the tilde character

    # Get all PDF files in the specified directory
    pdf_files = [os.path.join(pdf_directory, file) for file in os.listdir(pdf_directory) if file.endswith('.pdf')]

    if not pdf_files:
        print(f"No PDF files found in the directory: {pdf_directory}")
        return

    patient_info_list, result_summary_list = get_final_data_for_excel(pdf_files)
    output_excel_file = 'patient_details.xlsx'
    save_to_excel(output_excel_file, patient_info_list, result_summary_list)


if __name__ == "__main__":
    main()
