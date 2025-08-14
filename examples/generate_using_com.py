# coding=utf-8
import tempfile
from typing import List, Callable
from pymsword.docxcom_template import DocxComTemplate
from pymsword.com_utilities import table_inserter, heading_inserter, document_inserter

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
        "section":[
            {"header": heading_inserter("Section 1", 2),
             "content": "This is the content of section 1"},
            {"header": heading_inserter("Section 1.1", 3),
             "content": "This is the content of section 1.1"},
            {"header": heading_inserter("Section 2", 2),
             "content": "This is the content of section 2"}
        ],
        "rtf_content": document_inserter(os.path.join(os.path.dirname(__file__), "sample.rtf")),
    }
    #create temporary file
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as outfile:
        pass

    template.generate(data, outfile.name)
    print("Wrote temporary file", outfile.name)
    os.startfile(outfile.name)

if __name__=="__main__":
    main()