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

# Find all file sizes in ARCHIVE and store them in an array:

filesizes = [];
files_in_archive = os.listdir('ARCHIVE');

for file_in_dir in files_in_archive:
	if file_in_dir.endswith('.png'):
		filesizes.extend([os.path.getsize('ARCHIVE' + '\\' + file_in_dir)]);

# Iterate through TEMP and move unique files over to ARCHIVE

files_in_temp = os.listdir('TEMP');

for file_in_dir in files_in_temp:
	if file_in_dir.endswith('.png'):
		current_file_size = os.path.getsize('TEMP' + '\\' + file_in_dir);
		if current_file_size not in filesizes:
			# Copy the png-file
			shutil.copy('TEMP' + '\\' + file_in_dir,'ARCHIVE');
			# Copy the corresponding ppt-file:
			ppt_file_name = file_in_dir[:-4]+'.ppt';
			shutil.copy('TEMP' + '\\' + ppt_file_name,'ARCHIVE');
			# Put this file size in the array so that if we have more of this same file then it isn't moved over
			filesizes.extend([current_file_size]);

# Delete contents of TEMP

for root, dirs, files in os.walk('TEMP'):
    for f in files:
    	os.unlink(os.path.join(root, f));
    for d in dirs:
    	shutil.rmtree(os.path.join(root, d));