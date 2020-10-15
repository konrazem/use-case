# Tool to generate Use Case word document (docx) from Excel file 

## Installation python-docx and openpyxl

Program is based on 2 libraries:
1. [python-docx](https://python-docx.readthedocs.io/en/latest/) for docx files and 
2. [openpyxl](https://openpyxl.readthedocs.io/en/stable/) for xlsx or xls files.

Before you need to install those libraries with python 3.0. Note that python-docx need Visual Studio 2014 and relies on libxml2 and libxslt (lxml). **IF YOU DO NOT HAVE IT** then for Windows users go to [lxml](https://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml) and download specific whl file. For example for windows 64bit and python 3.9 install **lz4‑3.1.0‑cp39‑cp39‑win_amd64.whl**. Next install with pip for the given *wheel* in a way: 

```
pip install <PATH TO THE DOWNLOADED WHL FILE>.whl
```

Next you are good to install python-docx (on which this project depends on):

```
pip install python-docx
```

Next install [openpyxl](https://openpyxl.readthedocs.io/en/stable/#installation) library used to get data from excel:

```
pip install openpyxl
```

## Excel sheet (data.xlsx)

This is the file where you work on use cases. There are some rules for now for the excel file: 
1. Name of the file is data.xlsx
2. Names of the sheets are very important. For now, if you want to change the sheets names you need to change them also in [main.py](./main.py). Moreover there must be DESCRIPTION sheet for categories description.
3. The names of properties need to have 'Priority' and 'Name' fileds
4. First row with the fields names cannot contain empty values

## Diagram images (*.png)

If you want to put diagram images into document you need to name them same as categories to which they belong to. This forse a bit user to split diagrams for each category. Each image must be png extension.

## Tips

To show in heading status you can put in excel in Priority filed the information.