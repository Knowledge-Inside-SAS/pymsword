import os.path

from pymsword.docxcom_template import DocxComTemplate
from pymsword.com_utilities import document_inserter, image_inserter, update_document_toc
import win32com.client


class CachedExcelReader:
    def __init__(self):
        self.excel = None
        self.current_workbook = None
        self.current_document = None

    def _init_excel(self):
        if self.excel is None:
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False

    def open(self, document_path:str):
        """ Open an Excel document and cache the Excel application object.
        If the document is already open, it will return the cached workbook.
        """
        self._init_excel()
        document_path = os.path.abspath(document_path)
        if self.current_document != document_path:
            self._close()
            self.current_workbook = self.excel.Workbooks.Open(document_path)
            self.current_document = document_path

        return self.excel, self.current_workbook

    def _close(self):
        """ Close the current workbook and Excel application if they are open. """
        if self.current_workbook:
            self.current_workbook.Close(SaveChanges=False)
            self.current_workbook = None
            self.current_document = None

    def close(self):
        """ Close the Excel application and release resources. """
        self._close()
        if self.excel:
            self.excel.Quit()
            self.excel = None

def excel_sheet_inserter(cache:CachedExcelReader, excel_path:str, sheet_name:str):
    """Makes an inserter that inserts an Excel sheet into a Word document at the specified location.

    Args:
        cache (CachedExcelReader): Cached Excel reader instance.
        excel_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet to insert.

    Returns:
        DocumentInserter: An object that can be used to insert the Excel sheet into a Word document.
    """
    def insert(word_range):
        excel, workbook = cache.open(excel_path)
        try:
            sheet = workbook.Sheets(sheet_name)
            used_range = sheet.UsedRange
            used_range.Copy()  # copy to clipboard
            # paste it into the Word document at the specified range
            word_range.Paste()

            # Auto-fit all tables in the pasted range
            for table in word_range.Tables:
                table.AutoFitBehavior(2) # wdAutoFitWindow
            doc = word_range.Document
            word_range.Style = doc.Styles("Normal")
        finally:
            cache._close()  # Close the workbook but keep the Excel application open
    return insert


def main():
    template = DocxComTemplate("LAPTOP functional specification template.docx")
    source_dir = "data"


    data = {
        "doc_name": "Functional Specification"
    }
    functions = []
    data['function'] = functions
    #enumerate XSLX files in the source directory
    excel_cache = CachedExcelReader()
    try:
        for filename in sorted(os.listdir(source_dir)):
            if not filename.lower().endswith(".xlsx"): continue
            print("Processing file:", filename)
            func = {}
            functions.append(func)
            funcname = os.path.splitext(filename)[0]
            func['name'] = funcname
            excel_path = os.path.join(source_dir, filename)

            #func['presentation'] = "presentation"
            func['presentation'] = excel_sheet_inserter(
                excel_cache, excel_path, "Presentation"
            )
            func['specification'] = excel_sheet_inserter(
                excel_cache, excel_path, 2 #sheet index 2 corresponds to the second sheet
            )

            description_file = os.path.join(source_dir, funcname + ".rtf")
            if os.path.exists(description_file):
                func['description'] = document_inserter(description_file, set_style="Normal")

            diagram_file = os.path.join(source_dir, funcname + ".svg")
            if os.path.exists(diagram_file):
                func['diagram'] = image_inserter(diagram_file)

        template.generate(data, "LAPTOP functional specification result.docx",
                          postprocess=update_document_toc)
        os.startfile("LAPTOP functional specification result.docx")
    finally:
        excel_cache.close()

if __name__ == "__main__":
    main()
    # Make sure to close the Excel application when done
