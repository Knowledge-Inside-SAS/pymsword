# pymsword

**Pymsword** is a Python library for generating DOCX documents with simple templates.

---

## Features

- **Direct generation of DOCX documents with text and images** — provides good performance and reliability.
- **Jinja-like Placeholder Syntax** — easy-to-use templating for text and images.
- **COM Automation Support** — integrate COM calls to extend document generation capabilities, at the cost of performance.
- **MIT Licensed** — free and open-source with a permissive license.
---

## Template Syntax

Template syntax is inspired by Jinja, but only the small subset of features is implemented: 
- **Variables**: `{{ variable }}` for inserting variables.
- **Groups**: `{% group_name %} ... {% end group_name %} ` for grouping content.

### Variables
Use `{{variable}}` to insert values.
In the data, variables are defined as keys in a dictionary.

### Escaping Braces
To escape double braces, use: `{{"{{"}}`.

### Groups
Use `{% group_name %} ... {% end group_name %}` to define a group of content that can be repeated.
In the data, groups are defined as lists of dictionaries.

#### Extending group capture range
Due to the structure of DOCX documents, it is not possible to put text tags outside of the paragraph or table row.
Consequently, simple groups can not be used to generate table rows or lists.

To overcome this limitation, group's range can be extended using additional markers:

 - Use `{% group_name p %} ... {% group_name %}` for a group that captures the whole paragraph or list item.
 - Use `{% group_name row %} ... {% group_name %}` for a group that captures the whole table row.
 - Use `{% group_name cell %} ... {% group_name %}` for a group that captures the whole cell of a table. 

Example:
```python
template = DocxTemplate("extend_groups.docx")
data = {
    "items": [
    {"text": "Item 1"},
    {"text": "Item 2"},
    {"text": "Item 3"},
    ]
}
template.generate(data, "extend_groups_result.docx")
```
![Group example](/doc/capture.png)

## DOCX Document Generation
Pymsword supports two modes of document generation: direct DOCX generation, that is done using pure Python code.
It provides optimal performance and reliability, however it can only fill templates with plain text and images.

THe following example demonstrates how to use the library to generate a DOCX document with a template:
```python
from pymsword.docx_template import DocxTemplate, DocxImageInserter

# Load the template
template = DocxTemplate("template.docx")
# Define the data for the template
data = {
    "title": "My Document",
    "content": "This is a sample document.",
    "items": [
        {"name": "Item 1", "value": 10},
        {"name": "Item 2", "value": 20},
    ]
    "image": DocxImageInserter("image.png"),
}
# Render the template with data
template.generate(data, "output.docx")
```

The template and the resulting document are shown below:
![Template and result](/doc/template_basic.png)

For the more complete example refer to the `examples` directory.

### DOCX + COM Automation
More advanced documents can be generated using COM automation which gives access to the full range of Word features.
Its disadvantage is that it is significantly slower than direct DOCX generation, and requires Microsoft Word to be installed on the system.
To use COM automation, you need to install the `pywin32` package and use the `DocxComTemplate` class.

To insert data using COM, put *inserter fucntion* instead of the value in the data. Inserter functions
take single argument which is Word.Range object, and insert desired data into it.

Module `pymsword.com_utilities` provides some useful inserters.

```python
from pymsword.docxcom_template import DocxComTemplate
from pymsword.com_utilities import table_inserter

# Load the template
template = DocxComTemplate("template.docx")

# Define the data for the template
data = {
    "header": "My Document",
    "table": table_inserter([["Col1", "Col2"], ["Row1", "Row2"]])
}
# Render the template with data
template.generate(data, "output.docx")
```
The template and the resulting document are shown below:

![COM Template and result](/doc/template_com.png)

### Complete list of inserter functions, available in `pymsword.com_utilities` module:
| Function                               | Description                                                                                  |
|----------------------------------------|----------------------------------------------------------------------------------------------|
| `table_inserter(data:List[List[str]])  | Inserts a table with the given data. Each sublist represents a row.                          |
| `image_inserter(picture_path:str)`     | Inserts an image from the specified path using COM. Supports more formats than the DOCX mode |
| `document_inserter(document_path:str)` | Inserts content of another document, can be DOCX, RTF or anything supported by Word.       |
| `anchor_inserter(text:str, anchor:str)` | Inserts an anchor with the given text and name.                                           |
| `heading_inserter(text:str, level:int=1)` | Inserts a heading with the given text and level. Level 1 is the highest level.             |

## Requirements
- Python 3.7 or higher
- `pywin32` package for COM automation
- Microsoft Word installed (optional, for DOCX + COM mode)
- Pillow package for image handling 

## Installation

You can install Pymsword using pip:

```bash
pip install pymsword
```

## License
This project is licensed under the MIT License - see the [LICENSE.MIT](LICENSE.MIT) file for details.
