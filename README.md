# ImageMuncher
Makes a new powerpoint slide for each folder and
populates that slide with the given images and location specifications.

## Installation
If you have a Mac, you should already have pip.
Run `pip --version` from the console
if you aren't sure.
In the unlikely event you do not have pip,
you can download it [here](https://pip.pypa.io/en/stable/installing/).

This tentatively uses Python 3.*,
which is included if Anaconda3 or Miniconda3
(i.e the latest version) is installed. Alternatively, you could
install Python3 directly from their [website](https://www.python.org/downloads/).

Run `pip3 install python-pptx` from console to install the pptx module.


## Running the program
Navigate to the directory containing ImageMuncher.py in the console
and complete the following:

1. Initialize the variables `PATH_TO_DIRECTORIES` and `PATH_TO_POWERPOINT`
in `INPUTS.py` with the specified values.
2. Run the command `python3 ImageMuncher.py`.
If you cannot run python3, it means you may not have python3 installed.

_Recommended:
Move the ImageMuncher folder to the directory containing all of the subfolders.
Also, make sure there are no spaces in the name of the powerpoint._

_If the powerpoint of the specified name does not yet exist, one will be created._

## Warning
* This has only been run from a unix machine (mac) before,
so if you use a windows system... Use at your own risk?

<!-- INSERT PUBLICATION HERE AS WELL -->