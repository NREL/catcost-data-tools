# catcost-data-tools

Part of CatCost, a free catalyst cost estimator. For more info on CatCost, visit https://catcost.chemcatbio.org

This repository contains Python scripts to convert estimates from the Excel version of CatCost to JSON files compatible with the web version.

## Getting Started - GUI executable

Download an executable for your platform from [/releases/latest](https://github.com/NREL/catcost-data-tools/releases/latest)

## Getting Started - Python Scripting

1. Set up and activate a Python 3.8 virtual environment
2. Clone this repository and cd to the catcost-data-tools directory
3. Run `pip install -r requirements.txt`

### Running GUI from command line

With your Python 3.8 virtual environment activated,

`python gui.py`

### Building GUI executable

First install the development requirements into your Python 3.8 environment:

`pip install -r dev-requirements.txt`

Build for Windows:

`pyinstaller pyinstaller-win.spec`

Build for Mac:

`pyinstaller pyinstaller-mac.spec`
