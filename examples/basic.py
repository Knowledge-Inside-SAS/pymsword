from pymsword.docx_template import DocxTemplate, DocxImageInserter

# Load the template
template = DocxTemplate("basic.docx")
# Define the data for the template
data = {
    "title": "My Document",
    "content": "This is a sample document.",
    "items": [
        {"name": "Item 1", "value": 10},
        {"name": "Item 2", "value": 20},
    ],
    "image": DocxImageInserter("image.png"),
}
# Render the template with data
template.generate(data, "basic_result.docx")