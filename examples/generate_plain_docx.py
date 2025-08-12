# coding=utf-8
import tempfile

from pymsword.docx_template import DocxImageInserter, DocxTemplate

def main():
    import os
    template_path = os.path.join(os.path.dirname(__file__), "template.docx")
    image_path = os.path.join(os.path.dirname(__file__), "image.png")
    template = DocxTemplate(template_path)

    data = {
        "text_header": "Header text",
        "text": "Text in the document body",
        "image": DocxImageInserter(image_path),
        "list1": [
            {"value": "Item 1"},
            {"value": "Item 2"},
            {"value": "Item 3"},
        ],
        "list2": [
            {"value": "A"},
            {"value": "B"},
            {"value": "C"},
        ],
    }
    #create temporary file
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as outfile:
        pass
    template.generate(data, outfile.name)
    print("Wrote temporary file", outfile.name)
    os.startfile(outfile.name)

if __name__=="__main__":
    main()