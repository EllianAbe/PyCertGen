# PyCertGen
Looking to generate multiple certificates with just a single-line command? PyCertGen is the perfect solution for you.


## Installation
```shell
$ pip install pycertgen
```

## Usage
For gui usage:
```shell
python3 pycertgen --show_gui
```

for cli usage:

```
python3 pycertgen --certified_path="input/data.xlsx" --template_path="input/template.pptx" --key_autogen  --to_pdf_method="libreoffice" --output_path="output/"
```

## Templates consistency
You can create any template format based on your own needs. Follow the consistency rule:
Ensure that when creating or modifying any PowerPoint templates or Excel file, every column name from the Excel file must be represented as a placeholder in the [column_name] (column name between square brackets) format within the PowerPoint slides. This helps maintain consistency and clarity when populating the slides with data from the Excel columns.
