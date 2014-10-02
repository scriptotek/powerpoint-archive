These are some small scripts that look through powerpoint files and exports each slide as a .png screenshot and as it's own .ppt file.

### Usage

Run `run.bat` in a folder that has the folders TEMP and ARCHIVE within it.
Python must be installed, and you should edit `run.bat` to add the presentations to be archived and set paths.

### Description

* The javascript file `dump.js` creates individual ppt-files of each slide in the powerpoint-files in the array filearray.
It also takes a png-screenshot of each slide. These are moved to the folder TEMP.

* The python script `store.py` reads the TEMP-folder and moves unique presentations (i.e. slides) over to the ARCHIVE-folder.
Uniqueness is only determined by the file size.

* The bat-script `run.py` runs the javascript first and then the python script.
