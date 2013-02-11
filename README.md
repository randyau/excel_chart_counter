excel chart counter
===================

This is a very dumb one-off script that crawls through a list of directories, searches for excel files (specifically xlsx, but should handle xls fine w/ some tweaks to the regex). 

Then it goes through all the available worksheets and records the count of chart objects in the file so you have a dated series of charts created throughout a year.

There's some simple try blocks for it to keep going if it hits an exception while working with files, mostly this guards against badly named/temp files that can pop up.

Prerequisites
===
* A windows machine. COM and OS X probably don't work well together
* MS Excel that can handle the files you plan on chugging through.
* py-win32com
* py-dateutil


Input assumptions
===

IMPORTANT: This script assumes all files follow a naming convention of having YYYY-MM-DD\_filename.xlsx and it uses that as a the file's creation date. It does NOT read the file create/modify times.

It'll group all the data on a daily basis, going by the file name date string.

Output 
===
It'll spit the results of the count into results.txt, tab-deliminated, ready for dumping into excel to chart your chart of charts. 

The dates are unsorted, largely because of laziness.

Disclaimer
===
This is a toy script hacked together for fun and provided as-is. Use at your own risk. I'm not responsible if it destroys all your data.

