import re
import PyPDF2
import pandas as pd


def extract_patient_details(text):
    patient_name = get_patient_name(text)
    aml_ngs_panel = get_aml_ngs_panel(text)
    patient_id = get_patient_id(text)
    dob = get_dob(text)
    sex = get_sex(text)
    date_collected = get_date_collected(text)
    date_reported = get_date_reported(text)
    surg_path = get_surg_path(text)
    specimen_id = get_specimen_id(text)
    specimen_source = get_specimen_source(text)
    ordering_physician = get_ordering_physician(text)
    date_received = get_date_received(text)
    facility = get_facility(text)

    # Creating a dictionary to store extracted details
    patient_details = {
        'Patient Name': patient_name,
        'AML NGS Panel': aml_ngs_panel,
        'Patient ID': patient_id,
        'Date of Birth': dob,
        'Sex': sex,
        'Date Collected': date_collected,
        'Date Reported': date_reported,
        'Surg-Path #': surg_path,
        'Specimen ID': specimen_id,
        'Specimen Source': specimen_source,
        'Ordering Physician': ordering_physician,
        'Date Received': date_received,
        'Facility': facility
    }

    return patient_details


def process_pdf(pdf_path):
    text = extract_text_from_pdf(pdf_path)
    patient_details = extract_patient_details(text)
    patient_result_summary = extract_result_summary(text)
    clinical_relevant_results = extract_clinical_relevant_results(text)
    variants_of_uncertain_significance = extract_variants_of_uncertain_significance(text)
    technical_summary = extract_technical_summary(text)
    return patient_details, patient_result_summary, clinical_relevant_results, variants_of_uncertain_significance, technical_summary


def get_final_data_for_excel(pdf_files):
    patient_info_list = []
    result_summary_list = []
    clinical_relevant_results_list = []
    variants_of_uncertain_significance_list = []
    technical_summary_list = []

    # Iterate through PDF files
    for pdf_file in pdf_files:
        patient_info, result_summary, clinical_relevant_results, variants_of_uncertain_significance, technical_summary = process_pdf(
            pdf_file)
        # Append the patient information to the DataFrame
        patient_info_list.append(patient_info)
        result_summary_list.extend(result_summary)
        clinical_relevant_results_list.extend(clinical_relevant_results)
        variants_of_uncertain_significance_list.extend(variants_of_uncertain_significance)
        technical_summary_list.extend(technical_summary)
    return (patient_info_list, result_summary_list, clinical_relevant_results_list,
            variants_of_uncertain_significance_list, technical_summary_list)


def save_to_excel(output_excel_file, patient_info_list, result_summary_list, clinical_relevant_results_list,
                  variants_of_uncertain_significance_list, technical_summary):
    with pd.ExcelWriter(output_excel_file, engine='xlsxwriter') as writer:
        # Create DataFrame from the list
        patient_details_columns = ['Patient Name', 'AML NGS Panel', 'Patient ID', 'Date of Birth', 'Sex',
                                   'Date Collected', 'Date Reported',
                                   'Surg-Path #', 'Specimen ID', 'Specimen Source', 'Ordering Physician',
                                   'Date Received',
                                   'Facility']
        df_patient_details = pd.DataFrame(patient_info_list, columns=patient_details_columns)
        df_patient_details.to_excel(writer, sheet_name="PatientRecords")

        result_summary_columns = ["Specimen Id", "Date Collected", "Date Reported", "Variant Name", "p.",
                                  "NM_", "c.", "VAF"]
        df_result_summary = pd.DataFrame(result_summary_list, columns=result_summary_columns)
        df_result_summary.to_excel(writer, sheet_name="ResultSummary")

        clinical_relevant_results_columns = ["Specimen Id", "Date Collected", "Date Reported", "Variant Name", "p.",
                                             "NM_", "c.", "VAF"]
        df_clinical_relevant_results = pd.DataFrame(clinical_relevant_results_list,
                                                    columns=clinical_relevant_results_columns)
        df_clinical_relevant_results.to_excel(writer, sheet_name="ClinicalRelevantResults")

        variants_of_uncertain_significance_columns = ["Specimen Id", "Date Collected", "Date Reported", "Variant Name",
                                                      "p.",
                                                      "NM_", "c.", "chr", "VAF"]
        df_variants_of_uncertain_significance = pd.DataFrame(variants_of_uncertain_significance_list,
                                                             columns=variants_of_uncertain_significance_columns)
        df_variants_of_uncertain_significance.to_excel(writer, sheet_name="VariantsOfUncertainSignificance")

        technical_summary_columns = ["Specimen Id", "Date Collected", "Date Reported", "Variant Name", "p.",
                                     "c.", "chr", "Classification", "Coverage/VAF", "NM_"]
        df_technical_summary = pd.DataFrame(technical_summary,
                                            columns=technical_summary_columns)
        df_technical_summary.to_excel(writer, sheet_name="TechnicalSummary")

    print(f"Patient details written to {output_excel_file}")


def extract_text_from_pdf(pdf_path):
    text = ''
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
    return text


def extract_result_summary(text):
    pattern = re.compile(r'([A-Z0-9]+)\s*(p\.[^,]*),\s*(NM_[^,]+)\s*,\s*(c\.\s*.*)\s*VAF:\s*([^%]+%)')
    pattern_search = pattern.findall(text)
    result_summary = []
    for i in pattern_search:
        cleaned_match = tuple(part.replace('\n', '') for part in i)
        result_summary.append(cleaned_match)
    summary_list = []
    specimen_id = get_specimen_id(text)
    date_collected = get_date_collected(text)
    date_reported = get_date_reported(text)
    for summary in result_summary:
        summary_row = {
            "Specimen Id": specimen_id,
            "Date Collected": date_collected,
            "Date Reported": date_reported,
            "Variant Name": summary[0],
            "p.": summary[1],
            "NM_": summary[2],
            "c.": summary[3],
            "VAF": summary[4]
        }
        summary_list.append(summary_row)
    return summary_list


def extract_clinical_relevant_results(text):
    pattern = re.compile(
        r'([A-Z0-9]+)\s*(Interpretation:(.|\n)*?(?=p\.))?(p\.[^,\s]*)\s*(NM_[^,\s]+)\s*(c\.\s*.*)\s*VAF:\s*([^%]+%)')
    pattern_search = pattern.findall(text)
    clinical_relevant_results = []
    for i in pattern_search:
        cleaned_match = tuple(part.strip() for part in i)
        clinical_relevant_results.append(cleaned_match)
    clinical_relevant_results_list = []
    specimen_id = get_specimen_id(text)
    date_collected = get_date_collected(text)
    date_reported = get_date_reported(text)
    for summary in clinical_relevant_results:
        summary_row = {
            "Specimen Id": specimen_id,
            "Date Collected": date_collected,
            "Date Reported": date_reported,
            "Variant Name": summary[0],
            "p.": summary[3],
            "NM_": summary[4],
            "c.": summary[5],
            "VAF": summary[6]
        }
        clinical_relevant_results_list.append(summary_row)
    return clinical_relevant_results_list


def extract_variants_of_uncertain_significance(text):
    pattern = re.compile(r'VARIANTS\s*OF\s*UNCERTAIN\s*SIGNIFICANCE\s*((.|\n)*)\s*TECHNICAL\s*SUMMARY')
    pattern_search = pattern.findall(text)
    variants_of_uncertain_significance_text = ""
    if len(pattern_search) > 0:
        variants_of_uncertain_significance_text = pattern_search[0][0]
    pattern = re.compile(r'([A-Z0-9]+)\s*(p\..*)\s*(NM_.+)\s*:\s*(c\.\s*.*)\s*(chr.*)\s*VAF:\s*([^%]+%)')
    pattern_search = pattern.findall(variants_of_uncertain_significance_text)
    variants_of_uncertain_significance = []
    for i in pattern_search:
        cleaned_match = tuple(part.replace('\n', '') for part in i)
        variants_of_uncertain_significance.append(cleaned_match)
    variants_of_uncertain_significance_list = []
    specimen_id = get_specimen_id(text)
    date_collected = get_date_collected(text)
    date_reported = get_date_reported(text)
    for variant in variants_of_uncertain_significance:
        variant_row = {
            "Specimen Id": specimen_id,
            "Date Collected": date_collected,
            "Date Reported": date_reported,
            "Variant Name": variant[0],
            "p.": variant[1],
            "NM_": variant[2],
            "c.": variant[3],
            "chr": variant[4],
            "VAF": variant[5]
        }
        variants_of_uncertain_significance_list.append(variant_row)
    return variants_of_uncertain_significance_list

def extract_technical_summary(text):
    pattern = re.compile(r'TECHNICAL\s*SUMMARY\s*((.|\n)*)\s*')
    pattern_search = pattern.findall(text)
    technical_summary_text = ""
    if len(pattern_search) > 0:
        technical_summary_text = pattern_search[0][0]
    pattern = re.compile(
        r'([A-Z0-9]+)\s*(p\..*)\s*(c\.\s*.*)\s*(chr[^\s]+)\s*([A-Za-z\s]+)\s*(.*)\s*(NM_[0-9]+.[0-9]+)\s*')
    pattern_search = pattern.findall(technical_summary_text)
    technical_summary = []
    for i in pattern_search:
        cleaned_match = tuple(part.replace('\n', '') for part in i)
        cleaned_match = tuple(part.strip() for part in cleaned_match)
        technical_summary.append(cleaned_match)
    technical_summary_list = []
    specimen_id = get_specimen_id(text)
    date_collected = get_date_collected(text)
    date_reported = get_date_reported(text)
    for summary in technical_summary:
        variant_row = {
            "Specimen Id": specimen_id,
            "Date Collected": date_collected,
            "Date Reported": date_reported,
            "Variant Name": summary[0],
            "p.": summary[1],
            "c.": summary[2],
            "chr": summary[3],
            "Classification": summary[4],
            "Coverage/VAF": summary[5],
            "NM_": summary[6]
        }
        technical_summary_list.append(variant_row)
    return technical_summary_list


def get_match(text, pattern):
    match = re.search(pattern, text)
    return match.group(1).strip() if match else None


def get_specimen_id(text):
    patten = re.compile(r'Specimen\s*ID:\s+([^\n]+)\s+')
    return get_match(text, patten)


def get_aml_ngs_panel(text):
    pattern1 = re.compile(r'(Chronic\s*Myeloid\s*Neoplasm\s*Next\s*Generation\s*Sequencing\s*Panel)')
    pattern2 = re.compile(r'(Acute\s*Leukemia\s*Next\s*Generation\s*Sequencing\s*Panel)')
    pattern1_match = get_match(text, pattern1)
    if pattern1_match:
        return "No"
    pattern2_match = get_match(text, pattern2)
    if pattern2_match:
        return "Yes"


def get_patient_id(text):
    patten = re.compile(r'Patient\s*ID:\s+(.*?)\s+')
    return get_match(text, patten)


def get_patient_name(text):
    patten = re.compile(r'Name:\s+(.*?)\s+Surg')
    return get_match(text, patten)


def get_dob(text):
    patten = re.compile(r'DOB:\s+(.*?)\s+')
    return get_match(text, patten)


def get_sex(text):
    pattern = re.compile(r'Sex:\s+(.*?)\s+')
    return get_match(text, pattern)


def get_date_collected(text):
    pattern = re.compile(r'Date\s*Collected:\s+(.*?)\s+')
    return get_match(text, pattern)


def get_date_reported(text):
    pattern = re.compile(r'Date\s*Reported:\s+(.*?)\s+')
    return get_match(text, pattern)


def get_surg_path(text):
    pattern = re.compile(r'Surg.*Path #:\s+(.*?)\s+Patient')
    return get_match(text, pattern)


def get_specimen_source(text):
    pattern = re.compile(r'Specimen\s*Source:\s+([^\n]+)\s+')
    return get_match(text, pattern)


def get_ordering_physician(text):
    pattern = re.compile(r'Ordering\s*Physician:\s+(.*?)\s*Date\s*Collected:')
    return get_match(text, pattern)


def get_date_received(text):
    pattern = re.compile(r'Date\s*Received:\s+(.*?)\s+')
    return get_match(text, pattern)


def get_facility(text):
    pattern = re.compile(r'Facility:\s+([^\n]+)\s+')
    return get_match(text, pattern)
