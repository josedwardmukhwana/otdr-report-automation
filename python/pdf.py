import os
import xlwings as xw
import subprocess

class InvalidFormat(ValueError):
    """Custom exception for invalid date format."""
    pass

def pause_and_exit():
    input("Press any Enter to exit . . . ")
    exit()

def close_all_excel_apps():
    try:
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError as e:
        raise Exception("Failed to close Excel processes.")
        
def main():
    report_path = input("Enter report path: ")
    print('-----------------------------------------------------------------------------------------------------------------------')       

    if not os.path.exists(report_path):
        print("Invalid path. Please check the path and try again.")
        return
    print(f"Opening workbook: {report_path}")       

    try:
        workbook = xw.Book(report_path)
    except Exception as e:
        print(f"Error opening workbook: {e}")
        print('-----------------------------------------------------------------------------------------------------------------------') 
        return

    if not workbook.sheets:
        print("Workbook empty")
        print('-----------------------------------------------------------------------------------------------------------------------') 
        return

    num_sheets = len(workbook.sheets)
    print(f"Extracting a total of {num_sheets} reports")
    print('-----------------------------------------------------------------------------------------------------------------------') 

    pdfs_dir = os.path.join(os.path.dirname(report_path), "pdfs")
    os.makedirs(pdfs_dir, exist_ok=True)

    for sheet in workbook.sheets:
        pdf_path = os.path.join(pdfs_dir, f"{sheet.name}.pdf")
        sheet.api.ExportAsFixedFormat(0, pdf_path)
        print(f"Saved {sheet.name} as PDF: {pdf_path}")

    workbook.close()
    print('-----------------------------------------------------------------------------------------------------------------------') 
    close_all_excel_apps()
    pause_and_exit()

if __name__ == "__main__":
    main()
