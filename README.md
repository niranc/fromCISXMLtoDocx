# fromCISXMLtoDocx.py
## CIS Parser & Docx generator

<p>fromCISXMLtoDocx.py read CIS-CAT output XML file and write a report.</p>

<p>Please fill free to change the word template inside ./templates/template_word.doc</p>

<p>/!\ Works only on windows. /!\</p>

## Install
```sh
pip install -r requirements.txt
```

## How to use
Add your cis xml report into ./cis_xml/ and launch 
```
python fromCISXMLtoDocx.py --export-docx MyReport
```

## Usage
```sh
$ python fromCISXMLtoDocx.py  -h
Generate CIS report from XML output - V1.0.0

/!\     Please fill ./cis_xml/ with CIS-CAT XML output.  /!\

usage: fromCISXMLtoDocx.py [-h] [-v] [--export-docx EXPORT_DOCX]

A python script to generate CIS report.

optional arguments:
  -h, --help            show this help message and exit

Configuration:
  -v, --verbose         Verbosity level (-v for verbose, -vv for advanced, -vvv for debug)

Export:
  --export-docx EXPORT_DOCX
                        Output DOCX file to store the results in.
```

## TODO oneday
- Add CIS return command from XML
