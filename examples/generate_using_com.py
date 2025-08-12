# coding=utf-8
import tempfile
from typing import List, Callable
from pymsword.docxcom_template import DocxComTemplate


def table_inserter(data:List[List[str]])->Callable[[object], object]:
    """This function creates an "Inserter" - a function that takes Word.Range object
    And inserts a table into the document at the specified location
    """
    nrows = len(data)
    ncols = 0 if nrows == 0 else len(data[0])

    def inserter(word_range):
        #Insert a table into the document.
        if nrows == 0 or ncols == 0:
            return
        table = word_range.Document.Tables.Add(word_range,nrows,ncols,1,2) # wdWord9TableBehavior, wdAutoFitWindow
        for i, row in enumerate(data):
            for j, cell in enumerate(row):
                # Set the text of the cell
                table.Cell(i + 1, j + 1).Range.Text = str(cell)
        return
    return inserter


def main():
    import os
    template_path = os.path.join(os.path.dirname(__file__), "template_com.docx")
    template = DocxComTemplate(template_path)

    data = {
        "text": "This is a test of the docx template using COM",
        "table": table_inserter([
            ["Column 1", "Column 2", "Column 3"],
            ["A", "B", "C"],
            ["D", "E", "F"],
            ["G", "H", "I"]
        ]),
    }
    #create temporary file
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as outfile:
        pass
    template.generate(data, outfile.name)
    print("Wrote temporary file", outfile.name)
    os.startfile(outfile.name)

if __name__=="__main__":
    main()