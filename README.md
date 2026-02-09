# VOC_HTChar_Align.py

## Description

`VOC_HTChar_Align.py` is a script for inter-sample GC-MS data alignment and compound classification.

## What does the script do

The script takes `Excel` file as input and does the following:

1. Inter-sample alignment based on exact matching of the compound names and CAS# and tolerance-based (e.g., 0.1 min tolerance window) matching of retention times.
2. Retrieval of InChIKey and SMILES values from the PubChem database (Kim et al. 2023. Nucleic Acids Res.) based on the compound names.
3. Classification of the compounds using ClassyFire Batch by Fiehn Lab (https://cfb.fiehnlab.ucdavis.edu/) based on the PubChem InChIKey values.
4. Detection frequency calculation.

## Prerequisites

Before using the script, several applications/tools have to be installed:

1. Python 3 (https://www.python.org/downloads/windows/).
2. All necessary modules (see **lines 37-50**).
3. Google Chrome and ChromeDriver (https://developer.chrome.com/docs/chromedriver/downloads).

Modules can be intalled via Command Prompt as follows:

* Type `cmd` in the Search line >> Click on the Command Prompt icon >> Type: `pip install module_name` >> Press `Enter` >> Repeat for the next module.

Module names:

* tk
* pandas
* openpyxl
* pubchempy
* selenium

## How to use the script

The script should be executed in Command Prompt.

To use the script, the following steps must be executed:

1. Type `cmd` in the Search line >> Click on the Command Prompt icon >> Type: `python "the path to the script including the extension (.py)"`, e.g. `python "D:/Projects/VOC_HTChar_Align.py"` >> Press `Enter`.
2. Choose the files for processing in the new pop-up window and press `Open`. The processed file will be saved to the same location.

## Notes and recommendations

Before running the script, change the path to the working folder and the output file name (see **lines 434**).

This script is tailored to a certain data table format, i.e. the GCMSsolution (v4.52) software (Shimadzu, Japan) output file format.

The input file must contain the following columns to be processed:

* `"Substance"`, `"CAS"`, `"RT, min"`, `"Similarity, %"`, and `"Area"`.

## License

[![MIT License](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/license/mit)

Intended for academic and research use.
