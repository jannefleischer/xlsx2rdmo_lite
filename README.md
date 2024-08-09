# xlsx2rdmo_lite

This is a very limited tool to import a well built Excel-File into RDMO as a rough template for further work on it.

## Dependencies

- `pandas` is used to import the xlsx/ods into a dataframe. 
- `python-slugified` is used to create slugified names for attributes and keys.
- Also this uses `rdmo-client` (**in a customized version** till rdmo_client supports version 2.x of rdmo)
python>3.7

## Install

python -m pip install git+https://github.com/jannefleischer/xlsx2rdmo_lite.git

## Usage

```python
from xlsx2rdmo_lite import xlsx2rdmo_lite
importer = xlsx2rdmo_lite()
importer.init_rdmo_access(
    'https://your.deployment.example',
    token='sometoken' #or use auth=('user','password')
)
importer.import_to_rdmo(r"path/to/xlsxfile.xlsx")
```

## Limitations

- Only two languages: de and en are "supported".
- Only widgettype `text` is supported for now.
- Pages are identical with questionsets.
- Plenty others! - feel free to help out.

## Sample excel-file & file convention

You can find a sample xlsx-file in sample/sample.xlsx (though import of .ods should work too). Field names are mandatory. 

If you want to use answers in multiple catalogs (for extending later), keep the fields 1 and 2 and frage_de (at least 100 chars) identical, as thise fields are used to create attribute names.

## Todo

- replace pandas with something more lightweight
- 