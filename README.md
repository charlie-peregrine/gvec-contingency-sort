# PSSE Contingency Sorting

### Introduction
This program scrapes a folder of .con files with PSSE generated contingencies and filters them based on buses, duplicates, etc.

### Installation
1. Download zip file or clone repository.
2. Create a virtual environment and activate it (optional).
3. Install required packages with `pip install -r requirements.txt` or `python -m pip install -r requirements.txt` or `py -m pip install -r requirements.txt`
3. Edit config.ini to supply paths for input and output files and change flags for terminal ouput.
4. Run with `python main.py` or `py main.py`. If you created a separate config file you can instead run `python main.py myconfigfile.ini` or `py main.py myconfigfile.ini`
