from pymsword.docx_template import DocxTemplate, DocxImageInserter

# Load the template
template = DocxTemplate("extend_groups.docx")
# Define the data for the template
data = {
    "items": [
    {"text": "Item 1"},
    {"text": "Item 2"},
    {"text": "Item 3"},
    ]
}
# Render the template with data
template.generate(data, "extend_groups_result.docx")