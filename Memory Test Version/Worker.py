from PyQt5.QtCore import QThread, pyqtSignal
from multiprocessing import Pool
from ReadSheets import read_sheets
from ReadExcel import read_excel
from ntpath import basename
import pandas as pd
import itertools
import time
import sys
import os

def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]

class WorkerThread(QThread):
    
    def __init__(self, files, saveFile, outputFormat, delim, system = '', filename = '', sheets = False, threads = 4):
        super().__init__()     
        self.files = files
        self.saveFile = saveFile
        self.outputFormat = outputFormat
        self.delim = delim
        self.system = system
        self.filename = filename
        self.sheets = sheets
        self.threads = threads
    
    value = pyqtSignal(str)
    timerValue = pyqtSignal(int)
    startValue = pyqtSignal(int)
    endValue = pyqtSignal(int)
        
    def run(self):               
        
        try:      
            start = time.time()              
                   
            file_list = [filename for filename in self.files]
            
            self.value.emit("Importing Files:")
            
            for i in file_list:
                self.value.emit("     " + str(basename(i)))
             
            readstart = time.time()
            self.startValue.emit(35)
            
            with Pool(int(self.threads)) as pool:
                if self.sheets == True:
                    df_list= pool.map(read_excel, file_list)
                else:
                    df_list = list(itertools.chain.from_iterable(pool.map(read_sheets, file_list)))

        
            self.timerValue.emit(35)
            self.endValue.emit(1)
            readend = time.time()
            self.value.emit('Files Imported.')
            self.value.emit('Read time: ' + str(readend - readstart))
            loopstart = time.time()
            self.value.emit('Formatting Files...')
            self.timerValue.emit(35)
            self.timerValue.emit(0)
            self.startValue.emit(35)
            try:
                with Pool(int(self.threads)) as pool:
                    table = pool.map(self.system.formatting, df_list)  
            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print(exc_type, fname, exc_tb.tb_lineno)
                self.value.emit(str(e))
                self.endValue.emit(1)
                
            try:
                os.remove(basename('; '.join(self.filename)))
            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print(exc_type, fname, exc_tb.tb_lineno)
                self.value.emit(str(e))
                self.endValue.emit(1)
            
            self.timerValue.emit(35)
            self.endValue.emit(1)
            self.value.emit('Files Formatted.')
            loopend = time.time()
            self.value.emit('Loop time: ' + str(loopend - loopstart))            
            
            self.value.emit('Saving File...')
            
            dataFrames = []
            try:
                for item in table:
                    dataFrames.append(pd.DataFrame.from_dict(item, "index"))
            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print(exc_type, fname, exc_tb.tb_lineno)
                self.value.emit(str(e))
                self.endValue.emit(1)

                
            tableDf = pd.concat(dataFrames)
            
            writestart = time.time()    
                    
            if left(right(self.saveFile, 4),1) == '.':
                fileName = left(self.saveFile, len(self.saveFile) - 4)            
            else: 
                fileName = self.saveFile
            
            tableDf.to_csv(fileName +'.'+ self.outputFormat.lower(), sep = self.delim, index = False, chunksize = 100000, encoding='utf-8')
            
            writeend= time.time()
            self.value.emit('Save Complete.')
            self.value.emit('Write time: ' + str(writeend - writestart))
            self.value.emit("Formatting Complete.")
            end = time.time()
            
            self.value.emit("Time Taken: " + str(round(end - start, 2)) + " Seconds.")    
            
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
            self.value.emit(str(e))
            self.endValue.emit(1)