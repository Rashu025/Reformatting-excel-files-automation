
import pandas as pd
import glob, os
from os.path import basename
import shutil
from zipfile import ZipFile
from datetime import date
import sys
import logging

#log file
logging.basicConfig(filename='log_file.log', encoding='utf-8',level=logging.DEBUG,datefmt='%m-%d %H:%M',
                    format='%(asctime)s:%(levelname)s:%(name)s:%(message)s')

class File():
    def __init__(self,zipname):
        self.zipname =zipname
    def __repr__(self):
        return 'File(%r)' % (self.zipname)

# Returns the current local date
today = date.today()

# Sheet name array
sheet_name=[]

#User input
inputdir =os.getcwd()

#Make new folder
#dir is not keyword


# To get the csv at designated path
#for xls_file in glob.glob(os.path.join(inputdir,"*.xls*")):
#    data_xls = pd.read_excel(xls_file, index_col=None,sheet_name=None,engine='openpyxl')
#  
#    for key in data_xls: 
#         csv_file2 = os.path.splitext(xls_file)[0]+"-"+key+"-mpe"+".csv"
#         data_xls[key].to_csv(csv_file2)

for xls_file in glob.glob(os.path.join(inputdir,"*.xls*")):
   
    data_xls = pd.read_excel(xls_file, index_col=None,sheet_name=None,engine='openpyxl')

    for key in data_xls: 

        sheet_name.append(key)
        #print(sheet_name)

  
    for key in data_xls: 
         for index,value in zip(range(0,len(sheet_name)),sheet_name):
            if value == "Sheet1":
                indexfirst=index-1
         for index,value in zip(range(0,len(sheet_name)),sheet_name):
            if value == "Sheet2":
                indexsecond=index-1
         for index,value in zip(range(0,len(sheet_name)),sheet_name):
            if value == "Sheet3":
                indexthird=index-1
         for index,value in zip(range(0,len(sheet_name)),sheet_name):
            if value == "Sheet4":
                indexforth=index-1
         if (key=="Sheet1"):
             csv_file2 = os.path.splitext(xls_file)[0]+"-"+sheet_name[indexfirst]+"-mpe"+".csv" # for unique keys how to manage
             data_xls[key].to_csv(csv_file2,index=False)
             o = File(csv_file2)
             logging.debug('Desired csv file created sheet 1: %s',o)
             #logging.debug("Desired csv file created:",str(csv_file2))
         if (key=="Sheet2"):
             csv_file2 = os.path.splitext(xls_file)[0]+"-"+sheet_name[indexsecond]+"-mpe"+".csv"
             data_xls[key].to_csv(csv_file2,index=False)
             o = File(csv_file2)
             logging.debug('Desired csv file created sheet 2: %s',o)
         if (key=="Sheet3"):
             csv_file2 = os.path.splitext(xls_file)[0]+"-"+sheet_name[indexthird]+"-mpe"+".csv"
             data_xls[key].to_csv(csv_file2,index=False)
             o = File(csv_file2)
             logging.debug('Desired csv file created: %s',o)
         if (key=="Sheet4"):
             csv_file2 = os.path.splitext(xls_file)[0]+"-"+sheet_name[indexforth]+"-mpe"+".csv"
             data_xls[key].to_csv(csv_file2,index=False)
             o = File(csv_file2)
             logging.debug('Desired csv file created: %s',o)
         


# To create new folder for zip file
def makemydir(dirname):
  try:
    os.makedirs(dirname)
  except OSError:
    pass
  # let exception propagate if we just can't
  # cd into the specified directory
#  os.chdir(dirname)
CsvPath= inputdir+"\\Files"

makemydir(CsvPath)

ziploc=inputdir+"\\"+str(today)+"-New-mpe.zip"


o = File(ziploc)


# To zip the file 
f=ZipFile(str(today)+"-New-mpe.zip",'w')
for root,dires,files in os.walk(inputdir):
    for csv_files_zip in glob.glob(os.path.join(inputdir,"*.csv*")):
        f.write(csv_files_zip,os.path.relpath(csv_files_zip, root))
f.close()
logging.debug('Desired zip file created: %s',o)


#Remove csv
files_in_directory = os.listdir(inputdir)
filtered_files = [file for file in files_in_directory if file.endswith(".csv")]
for file in filtered_files:
    path_to_file = os.path.join(inputdir, file)
    os.remove(path_to_file)


try:
    shutil.move(ziploc, CsvPath)
except:
    logging.debug('Zip file already created')

#Remove zip
files_in_directory = os.listdir(inputdir)
filtered_files = [file for file in files_in_directory if file.endswith(".zip")]
for file in filtered_files:
    path_to_file = os.path.join(inputdir, file)
    os.remove(path_to_file)

#command = "print('m')"  #The command needs to be a string  serin send -p SQEComponent -s MPE -d ./all_files_in_this_folder/
#os.system(command)

sys.exit(0) 
