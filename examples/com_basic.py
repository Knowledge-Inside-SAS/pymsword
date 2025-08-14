from pymsword.docxcom_template import DocxComTemplate
from pymsword.com_utilities import table_inserter

# Load the template
template = DocxComTemplate("com_basic.docx")

# Define the data for the template
data = {
    "header": "My Document",
    "table": table_inserter([["Col1", "Col2"], ["Row1", "Row2"]])
}
# Render the template with data
template.generate(data, "com_basic_result.docx")

