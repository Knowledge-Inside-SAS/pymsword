# pymsword

**Pymsword** is a Python library for generating DOCX documents with simple templates.

---

## Features

- **Direct generation of DOCX documents with text and images** — provides good performance and reliability.
- **Jinja-like Placeholder Syntax** — easy-to-use templating for text and images.
- **COM Automation Support** — integrate COM calls to extend document generation capabilities, at the cost of performance.
- **MIT Licensed** — free and open-source with a permissive license.
---

## Template Syntax Overview

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

## Document Generation

The simplest example of generating a document without using COM automation:

```python
from pymsword.docx_template import DocxTemplate

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
}

# Render the template with data
template.generate(data, "output.docx")
```

Where `template.docx` contains:

```
{{title}}

{{content}}

{% items %}
Name: {{name}}, Value: {{value}}{% end items %}
```

For the more complete example refer to the `examples` directory.

### COM Automation
THe above example works by directly generating DOCX code

## License
This project is licensed under the MIT License - see the [LICENSE.MIT](LICENSE.MIT) file for details.
