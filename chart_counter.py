# 2013 Randy Au, a simple script that uses win32com to count charts in excel files

import os
import win32com
import win32com.client
import re
import datetime
import dateutil
import dateutil.parser

xl = win32com.client.Dispatch("Excel.Application")

#Turn off Excel's updating stuff so it doesn't flicker and waste tiem rendering
xl.Application.DisplayAlerts = False
xl.Application.ScreenUpdating = False
xl.Application.EnableEvents = False
xl.Application.AskToUpdateLinks = False 

dir_queue = ["C:\\Somedirectory\\with\\your"
"D:\\some\\other\\directory\\no\\trailing\\slashes\\pls"
]

YEAR_INTERESTED = '2012'

files_processed = []

final_filelist = []
#regex to grab dates for grouping
file_re =  re.compile("(\d{4}-\d{1,2}-\d{1,2})[-_].+\.xlsx")
for doc_root in dir_queue:
  filelist = os.listdir(doc_root)
  filelist = [x for x in filelist if "xlsx" in x]
  
  filelist = [x for x in filelist if file_re.search(x)] #only valid filenames past this point
  filelist2 = []
  for x in filelist:
    try: #skip any files that don't have YEAR_INTERESTED in the name
      if YEAR_INTERESTED in file_re.search(x).groups()[0]:
        filelist2.append(x)
    except:
      pass 
  final_filelist.append([doc_root, filelist2])

chart_history = {}

#now iterate through the whole of files
for doc_root, filelist in final_filelist:
  for fn in filelist:
    datekey = dateutil.parser.parse( file_re.search(fn).groups()[0] )
    charts = 0
    try:
      files_processed.append(doc_root+"\\"+fn)
      xl.WorkBooks.Open( doc_root+"\\"+fn )
      ws_len = len(xl.WorkBooks(fn).WorkSheets)
      for i in xrange(ws_len):
        #i+1 is 0indexed -> 1indexed
        charts += len(xl.WorkBooks(fn).WorkSheets(i+1).ChartObjects())
      xl.WorkBooks(fn).Close(SaveChanges=False)
    except:
      print fn,"failed"
    chart_history[datekey] = chart_history.setdefault(datekey,0)+charts
 
#List of files that were processed
dfile = open("debug.txt",'w')
dfile.write("\n".join(files_processed))
dfile.close()

#simple results output, tab separated
ofile = open("results.txt",'w')
for k,v in chart_history.items():
  ofile.write( k.date().isoformat() )
  ofile.write("\t")
  ofile.write( str(v) )
  ofile.write("\n")

#Reactivate all the updating stuff in excel, but it's still a bit sluggish after, might be easier to close the program if you want to actually use it
xl.Application.DisplayAlerts = True
xl.Application.ScreenUpdating = True
xl.Application.EnableEvents = True
xl.Application.AskToUpdateLinks = True
