These are some small scripts that look through powerpoint files and exports each slide as a .png screenshot and as it's own .ppt file.

### Usage

* Make sure Python (and Powerpoint) is installed.
* Edit `run.bat` to configure which presentations to archive (no config file yet).
  Also make sure the path at the top of the file points to the folder the script runs from.
* Use Task Scheduler to make `run.bat` run every night. Make sure to test that it actually runs OK.

### Description

* The javascript file `dump.js` creates individual ppt-files of each slide in the powerpoint-files in the array filearray.
It also takes a png-screenshot of each slide. These are moved to the folder TEMP.

* The python script `store.py` reads the TEMP-folder and moves unique presentations (i.e. slides) over to the ARCHIVE-folder.
Uniqueness is only determined by the file size.

* The bat-script `run.py` runs the javascript first and then the python script.
