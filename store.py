#!/usr/bin/env python
# -*- coding: utf-8 -*-

# script.py
# We have two folders: TEMP and ARCHIVE. The goal is to move files from TEMP
# to ARCHIVE, but only move the files we don't already have there. We determine
# the uniqueness of a file by file size only (should be sufficient for our
# purposes). The program will first get the sizes of all .png files in ARCHIVE
# and store them in an array. Then it look at all file sizes in TEMP, and if
# it finds a file size that isn't in the array then that file is moved over
# (.png and the corresponding .ppt).

import os, shutil;

from datetime import datetime
today = datetime.now().strftime('%Y-%m-%d')

# Set path to script folder
path_to_script = 'C:\\Users\\storskjerm\\Desktop\\powerpoint\\';
path_to_archive = path_to_script + 'ARCHIVE';
path_to_temp = path_to_script + 'TEMP';

# Find all file sizes in ARCHIVE and store them in an array:

filesizes = [];
files_in_archive = os.listdir(path_to_archive);

print 'Gathering filesize in archive dir...'
for file_in_dir in files_in_archive:
	if file_in_dir.endswith('.png'):
		filesizes.extend([os.path.getsize(path_to_archive + '\\' + file_in_dir)]);

# Iterate through TEMP and move unique files over to ARCHIVE

files_in_temp = os.listdir(path_to_temp);

print 'Gathering filesize in temp dir...'
newfiles = 0
for file_no, file_in_dir in enumerate(files_in_temp):
	if file_in_dir.endswith('.png'):
		current_file_size = os.path.getsize(path_to_temp + '\\' + file_in_dir);
		if current_file_size not in filesizes:
			newfiles += 1
			# Copy the png-file
			outname = '%s (%04d)' % (datetime.now().strftime('%Y-%m-%d'), file_no)
			shutil.copy(path_to_temp + '\\' + file_in_dir, path_to_archive + '\\' + outname + '.png');
			# Copy the corresponding ppt-file:
			ppt_file_name = file_in_dir[:-4]+'.ppt';
			shutil.copy(path_to_temp + '\\' + ppt_file_name,path_to_archive + '\\' + outname + '.ppt');
			# Put this file size in the array so that if we have more of this same file then it isn't moved over
			filesizes.extend([current_file_size]);


print '%s : Found %d new slide(s)' % (today, newfiles)
logfile = open('log.txt', 'a')
logfile.write('%s : Found %d new slide(s)\n' % (today, newfiles))

# n.strftime('%H:%M:%S')

# Delete contents of TEMP

for root, dirs, files in os.walk(path_to_temp):
    for f in files:
    	os.unlink(os.path.join(root, f));
    for d in dirs:
    	shutil.rmtree(os.path.join(root, d));