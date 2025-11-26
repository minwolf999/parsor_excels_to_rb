import os
from openpyxl import load_workbook

from excel_to_rb import ExcelToRb

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

if __name__ == "__main__":
    for excel in os.listdir(SCRIPT_DIR + '/excels'):
        if not excel.endswith('.xlsx'):
            if excel != '.gitkeep':
                print(f'{excel} ignored ! Please convert this file to xlsx format')
            continue
        
        excel_name = excel.replace('.xlsx', '')
        workbook_path = SCRIPT_DIR + '/excels/' + excel
        model_folder = SCRIPT_DIR + '/app/model/' + excel_name

        if not os.path.exists(model_folder):
            os.makedirs(model_folder)

        wb = load_workbook(filename=workbook_path, data_only=False)
        ExcelToRb(model_folder, wb, excel_name).exec()

        exit(0)