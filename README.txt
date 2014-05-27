These are some small scripts that look through powerpoint files and exports each slide as a .png screenshot and as it's own .ppt file.

HOW TO USE

For the script to work it has to be run in a folder that has the folders TEMP and ARCHIVE within it. Also, some information must be set in the javascript file.

- script.js
scriptpath = "C:\\Users\\Stian\\powerpoint\\"; // This is the path to where the script is located. Needed so that we can find the TEMP and ARCHIVE folders.
filearray = ['C:\\Users\\Stian\\powerpoint\\test.pptx','C:\\Users\\Stian\\powerpoint\\test2.pptx']; // Paths to all the powerpoint-files that should be a part of the archive.

WHAT THE DIFFERENT SCRIPTS DO

The javascript file script.js creates individual ppt-files of each slide in the powerpoint-files in the array filearray. It also takes a png-screenshot of each slide. These are moved to the folder TEMP.

The python script script.py reads the TEMP-folder and moved unique presentations (i.e. slides) over to the ARCHIVE-folder. Uniqueness is only determined by the file size.

The bat-script runs the javascript first and then the python script.