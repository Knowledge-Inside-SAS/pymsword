import os.path
from datetime import datetime

from pymsword.docxcom_template import DocxComTemplate
from pymsword.com_utilities import document_inserter
import win32com.client


def chart_inserter(workbook, chart_name:str):
    """Make inserter, that inserts given Excel chart
    :workbook: Excel.Workbook instance
    :
    """
    def insert(word_range):
        # Get the chart by name
        print("Inserting chart", chart_name)
        chart = workbook.Sheets(1).ChartObjects(chart_name)
        print("Copying chart", chart_name, chart)
        # Copy the chart
        chart.Copy()
        print("Chart copied")
        # Paste it into the Word document
        word_range.Paste()
        print("Chart pasted")
    return insert


def main():
    #Open the Excel file, using COM (because we need to extract diagrams from it)
    excel = win32com.client.Dispatch("Excel.Application")
    source = excel.Workbooks.Open(os.path.abspath("Sales sheet.xlsx"))

    try:
        sheet = source.Worksheets("sales_csv")

        max_row = sheet.UsedRange.Rows.Count
        table_data = {}
        for row in range(48, max_row+1):
            key = sheet.Cells(row, 4).Value
            if key is None:
                continue
            table_data[key] = sheet.Cells(row, 5).Value

        data = {
            "TOTAL_REVENUE": "{:,.2f}".format(table_data["Total Revenue"]),
            "GROWTH_PERCENT": "{:.0f}%".format(table_data["Groth/loss versus Q-1"] * 100),
            "TOP_REGION": table_data["Best performing region"],
            "TOP_PRODUCT": table_data["Top Product"],
            "INSERT_MONTHLY_CHART": chart_inserter(source, "Chart 1"),
            "INSERT_REGIONAL_CHART": chart_inserter(source, "Chart 2"),
            "INSERT_PRODUCT_CHART": chart_inserter(source, "Chart 3"),
            "REGIONAL_RECOMMENDATION": table_data["Regional strategy"],
            "PRODUCT_RECOMMENDATION": table_data["Product focus"].replace("\n", ", "),
            "GENERATION_DATE": datetime.now().strftime("%Y-%m-%d"),
            "DOC_ID": "00001"
        }

        template = DocxComTemplate(os.path.join(os.path.dirname(__file__), "QUARTERLY SALES REPORT template.docx"))

        template.generate(data, "quarterly sales report.docx")
    finally:
        # Close the Excel file
        source.Close(False)
        # Quit Excel application
        excel.Quit()
    os.startfile("quarterly sales report.docx")

if __name__ == "__main__":
    main()