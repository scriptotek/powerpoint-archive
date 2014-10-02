@echo off

cd C:\Users\storskjerm\Desktop\powerpoint

:: This script will make one-page powerpoint presentations
:: and take screenshots of each slide and put eveything in TEMP
cscript //nologo dump.js "\\\\ubreal54\\SHOW\\script\\_active.ppt"
cscript //nologo dump.js "\\\\ubreal54\\SHOW\\script\\_active.pptx"
timeout 3
cscript //nologo dump.js "\\\\ubreal60\\SHOW\\script\\_active.ppt"
cscript //nologo dump.js "\\\\ubreal60\\SHOW\\script\\_active.pptx"
timeout 3
cscript //nologo dump.js "\\\\ubreal42\\SHOW\\script\\_active.ppt"
cscript //nologo dump.js "\\\\ubreal42\\SHOW\\script\\_active.pptx"
timeout 3
cscript //nologo dump.js "\\\\ubreal36\\SHOW\\script\\_active.ppt"
cscript //nologo dump.js "\\\\ubreal36\\SHOW\\script\\_active.pptx"
timeout 3
cscript //nologo dump.js "\\\\ubreal41\\SHOW\\script\\_active.ppt"
cscript //nologo dump.js "\\\\ubreal41\\SHOW\\script\\_active.pptx"
timeout 3


:: This call will move unique files from TEMP to ARCHIVE, and then empty TEMP
python script.py

:: pause