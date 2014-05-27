@echo off

:: This script will make one-page powerpoint presentations and take screenshots of each slide and put eveything in TEMP
cscript script.js

:: This call will move unique files from TEMP to ARCHIVE, and then empty TEMP
python script.py

pause