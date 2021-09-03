import analysis
import master

FOLDER_ROOT_LOCATION = "/Users/wxp/Downloads/PHASE ONE CODED"

if __name__ == '__main__':
    # Merge to a file
    output_path, records_count = master.merge_excel_sheet(FOLDER_ROOT_LOCATION)

    # Create analysis sheet
    analysis.create_analysis_sheet(output_path, records_count)
