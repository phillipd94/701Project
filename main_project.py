# -*- coding: utf-8 -*-
"""
Created on Tue Nov 22 13:43:17 2016

@author: Phillip Dix
"""

#should probably figure out how unit tests work and add some to this thing
#should probably put some docstrings and stuff in here too

from __future__ import division
import traceback
import sys
import numpy as np
import os
import re
import openpyxl as xl
import encodeindex
import pickle as pickle

#from linetextedit1 import LineTextWidget

from PyQt5.QtWidgets import (QTabWidget,QApplication, QDialog, QLineEdit, QVBoxLayout,QHBoxLayout, QGridLayout, QTableWidget, QTableWidgetItem,
                             QMainWindow, QAction, QFileDialog, QMessageBox, QComboBox, QTextEdit, QPushButton,QHeaderView)
#from PyQt5.QtGui import QColor, QTextCursor,QBrush,QTextCharFormat
import PyQt5.QtGui as QtGui
from PyQt5.QtCore import QObject, pyqtSignal, Qt,QRegExp

from PyQt5.Qt import QFrame, QWidget, QHBoxLayout, QPainter
 

class EmittingStream(QObject): #http://stackoverflow.com/questions/8356336/how-to-capture-output-of-pythons-interpreter-and-show-in-a-text-widget

    textWritten = pyqtSignal(str)

    def write(self, text):
        try:
            self.textWritten.emit(str(text))
        except:
            msg=QMessageBox()
            msg.setText('error: Phillip is bad at code')
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()

    


class MainWindow(QMainWindow):
    
    def __init__(self, parent=None) :
        super(MainWindow, self).__init__(parent)
        
        
        #Change the output stream
        
    
        # Create the file menu
        
        self.menuFile = self.menuBar().addMenu("&File")
        self.actionSaveAs = QAction("&Save Spreadsheet As", self)
        self.actionSaveAs.triggered.connect(self.saveas)
        
        self.actionSSaveAs = QAction("&Save Script As", self)
        self.actionSSaveAs.triggered.connect(self.ssaveas)
        
        self.actionclosetab = QAction("&Close Current WorkBook", self)
        self.actionclosetab.triggered.connect(self.closewb)
        
        self.actionSOpen = QAction("&Open Script", self)
        self.actionSOpen.triggered.connect(self.openscript)
        ########SAVE THIS IT'S CLOSE TO WORKING
#        self.actionVSaveAs = QAction("&Save Variable Space As", self)
#        self.actionVSaveAs.triggered.connect(self.vsaveas)
#        
#        self.actionOpenvars = QAction("&Open Variable Space", self)
#        self.actionOpenvars.triggered.connect(self.openvars)
        #############################################
        self.actionNewSpreadsheet = QAction("&New Spreadsheet", self)
        self.actionNewSpreadsheet.triggered.connect(self.newspreadsheet)
        
        self.actionOpenSpreadsheet = QAction("&Open Spreadsheet", self)
        self.actionOpenSpreadsheet.triggered.connect(self.openspreadsheet)
        
        self.actionQuit = QAction("&Quit", self)
        self.actionQuit.triggered.connect(self.close)#self.close is predefined, so I didn't have to change anything here
        self.menuFile.addActions([self.actionNewSpreadsheet,self.actionOpenSpreadsheet,self.actionSaveAs,self.actionSSaveAs,self.actionSOpen,self.actionclosetab, self.actionQuit])
        
        # Set the central widget
        self.widget1 = Form()
        self.setCentralWidget(self.widget1)
        
        sys.stdout = EmittingStream(textWritten=self.normalOutputWritten)
        
        

#        sys.stderr = EmittingStream(textWritten=self.normalOutputWritten)
        self.newspreadsheet()
        
                
#        try:
#        self.widget1.command.returnPressed.connect(self.widget1.userExec)
#        except:
#            print "error connection failed"



########SAVE THIS IT'S CLOSE TO WORKING
#    def vsaveas(self):
#        fname = unicode(QFileDialog.getSaveFileName(self, "Save Variable Space as...")[0])
#        if fname:
#            b=open(fname,'w')
#            pickle.dump(self.widget1.locals,b)
#            b.close()
#
#    def openvars(self):
#        fname = unicode(QFileDialog.getOpenFileName(self, "Open")[0])
#        try:
#            b=open(fname,'r')
#            pickle.load(self.widget1.locals,b)
#            b.close()
#        except:
#            print "error, bad file type"
#################
    def closewb(self):
        del self.widget1.scr.edit.activeWB[self.widget1.scr.edit.WBindex]
        del self.widget1.scr.edit.activeWS[self.widget1.scr.edit.WBindex]
        self.widget1.table.removeTab(self.widget1.scr.edit.WBindex)

    def ssaveas(self):
        fname = unicode(QFileDialog.getSaveFileName(self, "Save Script as...")[0])  #AT LAST I HAVE MY VICTORY, THIS METHOD RETURNS A TUPLE, AND I HAD TO SELECT A SPECIFIC VALUE TO PASS TO THE UNICODE TYPE CAST
        if fname :
            try:
                 a=self.widget1.scr.edit.toPlainText()
                 b=open(fname+".py","w")
                 b.write(a)
                 b.close
            except:
                raise ValueError    
                
    def openscript(self):
        fname = unicode(QFileDialog.getOpenFileName(self, "Open")[0])
        if '.py' in fname or '.txt' in fname:
            a=open(fname,'r')
            b=a.read()
            self.widget1.scr.edit.setPlainText(b)
        
    def newspreadsheet(self):
#        self.widget1.scr.edit.activeWB.append(xl.Workbook())
        self.widget1.scr.edit.activeWB.append(xl.Workbook())
        d=self.widget1.scr.edit.activeWB[len(self.widget1.scr.edit.activeWB)-1].worksheets
        d[0]["A1"]=0
#        d=self.widget1.scr.edit.activeWB.worksheets
        self.widget1.scr.edit.activeWS.append(d)
#        print type(self.widget1.scr.edit.activeWS[0])
#        self.widget1.scr.edit.activeWS=d
        z=QTableWidget(0,0,self)
        v=QTabWidget()
        v.currentChanged.connect(self.widget1.switch_tabs_ws)
        s=self.widget1.scr.edit.activeWS[0][0]
        for i in s['A1':'Z100']:
                for j in i:
                    j.value=''
        v.addTab(z,"NewSheet")
        self.widget1.table.addTab(v,"Book1")
        self.widget1.updateUI()
        
        
    def openspreadsheet(self):
        fname = unicode(QFileDialog.getOpenFileName(self, "Open")[0])
        if '.xlsx' in fname:
            self.widget1.scr.edit.WBNames.append(fname)
            self.widget1.scr.edit.activeWB.append(xl.load_workbook(fname))
            d=self.widget1.scr.edit.activeWB[len(self.widget1.scr.edit.activeWB)-1].worksheets
            self.widget1.scr.edit.activeWS.append(d)
#            d=self.widget1.scr.edit.activeWB.worksheets
#            self.widget1.scr.edit.activeWS.append(d)
#     #       self.widget1.scr.edit.activeWS.append(self.widget1.scr.edit.activeWB[len(self.widget1.scr.edit.activeWB)-1].active)
            v=QTabWidget()
            for m in range(0,len(d)):
                z=QTableWidget(0,0,self)
                z.setRowCount(d[m].max_row)
                z.setColumnCount(d[m].max_column)
                a=d[m].iter_rows()
                for i in range(0,int(d[m].max_row)):
                    try:
                        b=a.next()
                        for j in range(0,int(d[m].max_column)):
                            if (str(b[j].value) != 'None'):
                                z.setItem(i,j,QTableWidgetItem(str(b[j].value)))
                    except:
                        pass
                v.addTab(z,str(d[m].title))
                v.currentChanged.connect(self.widget1.switch_tabs_ws)
            self.widget1.table.addTab(v,str(fname))
        else:
            msg=QMessageBox()
            msg.setText('error: bad file type')
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()
        
    def saveas(self) :
        """ Save the computed data as a text file. """


        fname = unicode(QFileDialog.getSaveFileName(self, "Save as...")[0])  #AT LAST I HAVE MY VICTORY, THIS METHOD RETURNS A TUPLE, AND I HAD TO SELECT A SPECIFIC VALUE TO PASS TO THE UNICODE TYPE CAST
        if fname :
        


            try:

                 self.widget1.scr.edit.activeWB.save(filename = fname)
                
            except:
                raise ValueError
                
    def normalOutputWritten(self, text):
        """Append text to the QTextEdit."""

#        cursor = self.widget1.cmd.textCursor()
#        cursor.movePosition(QTextCursor.End)
#        cursor.insertText(text)
#        self.widget1.cmd.setTextCursor(cursor)
#        self.widget1.cmd.ensureCursorVisible()
        if text!="\n":
            self.widget1.cmd.append("\n"+text)
                
class Form(QDialog) :

    def __init__(self, parent=None) :
        super(Form, self).__init__(parent)

        # Define three text boxes.
        self.scr=LineTextWidget()
#        self.scr = script_box()
        self.scr.edit.setPlainText("for i in range(0,5):\n\ta='output_test_'+str(i)\n\tb=open(a,'w')\n\tb.write(a)\n\tb.close\n")
        self.run_button = QPushButton('Run', self)
        self.clear_button = QPushButton('Clear Variables',self)
        self.clear_button.clicked.connect(self.clear_vars)
        self.cmd=outputbox()
        
        self.command=commandLine()
        try:
            self.command.returnPressed.connect(self.userExec) 
        except:
            print "error-yousuck"
            
        self.table = QTabWidget(self)
        
        self.vars = QTableWidget(5,2,self)

#        header.setSectionResizeMode(1)
#        header.setSectionResizeMode(3)


        header = self.vars.horizontalHeader()
        header.setSectionResizeMode(3)
        header.setStretchLastSection(True)

#        self.l=LineTextWidget()

        # Define the layout
        layout = QGridLayout()
        layout1=QVBoxLayout()
        layout1.addWidget(self.run_button)
        layout1.addWidget(self.scr)
        layout1.addWidget(self.command)
        layout1.addWidget(self.cmd)
        layout2=QHBoxLayout()
        layout2.addLayout(layout1)
        layout3=QVBoxLayout()
        layout3.addWidget(self.clear_button)
        layout3.addWidget(self.vars)
        layout2.addLayout(layout3)
#        layout1.addWidget(self.l)
        layout.addWidget(self.table,0,0)
#        layout.addWidget(self.vars,0,2)
        layout.addLayout(layout2,0,1)
        self.setLayout(layout)
        self.table.currentChanged.connect(self.switch_tabs_wb)
        # Connect our output box to the update function
        try:
            self.run_button.clicked.connect(self.userExec) 
        except:
            print "error"
        self.locals={}
        self.flag=True

    def clear_vars(self):
        
        self.locals=()
        self.flag=True
        self.vars.setColumnCount(2)
        self.vars.setRowCount(5)
        for i in range(0,self.vars.columnCount()):
            for j in range(0,self.vars.rowCount()):
                self.vars.setItem(i,j,QTableWidgetItem(""))

        
    def userExec(self) :
        """ Method for executing user code"""
        signal_sender=self.sender()
        fil_g5h6j1n2l530kdmcgkr=open("script.py",'w')#anything I declare here 
#will interfere with the user input when it is imported, hence the name
        fil_g5h6j1n2l530kdmcgkr.write(self.scr.edit.getCode())
        fil_g5h6j1n2l530kdmcgkr.close()
#        a=a.replace(" ","")
#        self.scr.setHtml(a)
#        b=os.path.abspath("mydir/myfile.txt")
        cmd=0
        if isinstance(signal_sender, commandLine):
            cmd=15

#        os.system('python script.py')#this is one possible way to do it
#        import script
        flag_error_exception=True
        try:
            if self.flag:
                if cmd==0:
                    exec(self.scr.edit.getCode())
                    self.flag=False
                elif cmd == 15:
                    exec(self.command.getCode())
                    self.flag=False
                else:
                    print "error in execution handling:1"
            else:
                if cmd==0:
                    glob=globals()
                    exec(self.scr.edit.getCode(),glob,self.locals)  #this is mind boggling
    #            print self.locals
                elif cmd==15:
                    glob=globals()
                    exec(self.command.getCode(),glob,self.locals)
                else:
                    print "error in execution handling:2"
        except SyntaxError as err:  #http://stackoverflow.com/questions/28836078/how-to-get-the-line-number-of-an-error-from-exec-or-execfile-in-python
            flag_error_exception=False
            error_class = err.__class__.__name__
            detail = err.args[0]
            line_number = err.lineno
        except Exception as err:
            flag_error_exception=False
            error_class = err.__class__.__name__
            detail = err.args[0]
            cl, exc, tb = sys.exc_info()
            line_number = traceback.extract_tb(tb)[-1][1]
        else:
            pass
        if not flag_error_exception:
            if cmd==0:
                print("%s at line %d: %s" % (error_class, line_number-2, detail))
            else:
                print("%s: %s" % (error_class, detail))



#        exec(self.scr.getCode())

        
#        try:
        self.updateUI()
        count=0
        self.vars.setColumnCount(2)
        self.vars.setRowCount(len(self.locals)-6)
#        print self.locals
        for i in self.locals:
#            if (i!="wb")| (i!="cmd")| (i!="flag_error_exception") | (i!="self")|(i!="fil_g5h6j1n2l530kdmcgkr"):
            if (i!="wb") and (i!="cmd") and (i!="flag_error_exception") and (i!="self") and (i!="fil_g5h6j1n2l530kdmcgkr") and (i!="signal_sender"):
                
                self.vars.setItem(count,0,QTableWidgetItem(str(i)))
#                print str(i)+"\n"
                self.vars.setItem(count,1,QTableWidgetItem(str(self.locals[i])))
#                print str(self.locals[i])+"\n"
                count+=1
#        except:
#            print "error updating UI"

    def addWS(self,a):
        self.scr.edit.activeWB[self.scr.edit.WBindex].create_sheet(a)
        self.scr.edit.activeWS[self.scr.edit.WBindex].append(self.scr.edit.activeWB[self.scr.edit.WBindex].worksheets[-1])
        return self.scr.edit.activeWS
        
    def addWB(self,a):
#        a=a.group()
        self.scr.edit.activeWB.append(xl.Workbook())
#        d=self.scr.edit.activeWB[len(self.scr.edit.activeWB)-1].worksheets
        d=self.scr.edit.activeWB[-1].worksheets
#        d=self.widget1.scr.edit.activeWB.worksheets
        self.scr.edit.activeWS.append(d)
#        print type(self.widget1.scr.edit.activeWS[0])
#        self.widget1.scr.edit.activeWS=d
        z=QTableWidget(0,0,self)
        v=QTabWidget()
        v.currentChanged.connect(self.switch_tabs_ws)
        s=self.scr.edit.activeWS[-1][0]
        for i in s['A1':'Z100']:
                for j in i:
                    j.value=''
        v.addTab(z,a)
        self.table.addTab(v,a)
        return self.scr.edit.activeWS
        
#    def setActiveWS(self,a):
        
        
    def updateUI(self):
        k=self.scr.edit.activeWB        
        for s in range(0,len(k)):
     
            d=self.scr.edit.activeWB[s].worksheets
#            print d
            for i in range(0,len(d)-self.table.widget(s).count()):
                c=self.table.widget(s).count()
                z=QTableWidget(1,1,self)
                if(d[c]["A1"]==None):
                    d[c]["A1"]=0
                self.scr.edit.activeWS[s].append(d[c])
#                z.currentChanged.connect(self.switch_tabs_ws)
                self.table.widget(s).addTab(z,str(d[c].title))
    #        self.widget1.scr.edit.activeWS.append(d)
    #            self.widget1.scr.edit.activeWS.append(self.widget1.scr.edit.activeWB[len(self.widget1.scr.edit.activeWB)-1].active)
            for v in range(0,len(d)):
                z=self.table.widget(s).widget(v)
                z.setRowCount(d[v].max_row)
                z.setColumnCount(d[v].max_column)
                a=d[v].iter_rows()
                for i in range(0,int(d[v].max_row)):
                        b=a.next()
#                        print z
                        for j in range(0,int(d[v].max_column)):
                            if (str(b[j].value) != 'None'):
                                z.setItem(i,j,QTableWidgetItem(str(b[j].value)))
#                                print str(b[j].value)

#            self.scr.edit.activeWS=self.scr.edit.activeWB.active
#            self.table.setRowCount(self.scr.edit.activeWS.max_row)
#            self.table.setColumnCount(self.scr.edit.activeWS.max_column)
#            a=self.scr.edit.activeWS.iter_rows()
#            for i in range(0,int(self.scr.edit.activeWS.max_row)):
#                b=a.next()
#                for j in range(0,int(self.scr.edit.activeWS.max_column)):
#                    if (str(b[j].value) != 'None'):
#                        self.table.setItem(i,j,QTableWidgetItem(str(b[j].value)))
#                    else:
#                        self.table.setItem(i,j,QTableWidgetItem(str("")))
    def switch_tabs_ws(self):
        try:
            self.scr.edit.WSindex=self.table.widget(self.scr.edit.WBindex).currentIndex()
        except AttributeError: #this is called when the tab is first initiallized but before there's a workbook and it causes errors
            pass

    def switch_tabs_wb(self):
        self.scr.edit.WBindex=self.table.currentIndex()
        self.switch_tabs_ws()
#        print self.scr.edit.WBindex
        
class script_box(QTextEdit):

    def __init__(self, parent=None) :
        super(script_box, self).__init__(parent) 
#        self.setAcceptRichText(False)
        self.activeWB=[]
        self.activeWS=[]
        self.WBindex=0
        self.WSindex=0
        self.WBNames=[]
        self.setTabStopWidth(50)
        self.textChanged.connect(self.updateScript)

    def updateScript(self):
##        try:
###            self.setHtml(self.getInput())
##            a=self.getInput()
##            fil_g5h6j1n2l530kdmcgkr=open("html.py",'w')#anything I declare here 
###will interfere with the user input when it is imported, hence the name
##            fil_g5h6j1n2l530kdmcgkr.write(a)
##            fil_g5h6j1n2l530kdmcgkr.close()
##        except:
##            msg=QMessageBox()
##            msg.setText('error: Phillip Dix is bad at code')
##            msg.setStandardButtons(QMessageBox.Ok)
##            msg.exec_()
#
#        a=re.finditer(r'[ \t\n\r\f\v+-=*()][A-Z]+\d+[ \t\n\r\f\v+-=*()]',self.document().toPlainText())
#        while (1==1):
#            try:
#                b=a.next()
#                c=b.group()
##                print c
#                self.find(c)
#                x=self.currentCharFormat()
#                x.setFontWeight(150)
#                self.setCurrentCharFormat(x)
##                cursor = self.textCursor()
##
##                format = cursor.charFormat()
##
##                format.setBackground(Qt.red)
##                format.setForeground(Qt.blue)
##                textSelected = cursor.selectedText()
##                print textSelected
##                cursor.setCharFormat(format)
###                self.setTextColor(QColor(0, 0, 255, 127))
#            except StopIteration:
#                break
##            except:  #commented for report writing
##                pass
##                msg=QMessageBox()
##                msg.setText('error: Phillip Dix is bad at code')
##                msg.setStandardButtons(QMessageBox.Ok)
##                msg.exec_()
        
        tempcurs=self.textCursor()
        a=tempcurs.position()
#        print a
        script=self.getInput().replace('<span style=" color:#0000ff;">',"")#FUCK HTML AND CSS
        


        fil_g5h6j1n2l530kdmcgkr=open("html.html",'w')#anything I declare here 
#will interfere with the user input when it is imported, hence the name
        fil_g5h6j1n2l530kdmcgkr.write(script)
        fil_g5h6j1n2l530kdmcgkr.close()

        
        self.blockSignals(True) #block signals to prevent infinite loops
        self.document().setHtml(script)
        self.blockSignals(False)
        tempcurs.setPosition(a)
        self.setTextCursor(tempcurs)

    def getCode(self):
        #method for parsing and returning user input code
        script=self.document().toPlainText()
        script='\n'+script+'\n'
        pattern0=r'[ \t\n\r\f\v+-=*()\[\]][A-Z]+\d+[ \t\n\r\f\v+-=*()\.]'
        pattern1=r'[ \t\n\r\f\v+-=*()\[][A-Z]+\d+[ \t\n\r\f\v+-=*()\.]'
        pattern2=r'=(.*wb\[self.scr.edit.WBindex\]\[self.scr.edit.WSindex\]\["[A-Z]+\d+"\].*)+'
        pattern3=r'\[\s*[A-Z]+\d+\s*-\s*[A-Z]+\d+\s*\]'
        pattern4=r'\snewWorkbook\(["\'].+["\']\)\s'
        pattern5=r'\ssetActiveWorkbook\(\d+\)\s'
        pattern6=r'\snewWorksheet\(["\'].+["\']\)\s'
        pattern7=r'\ssetActiveWorksheet\(\d+\)\s'
        re.MULTILINE=True
        script=re.sub(pattern3,self.rep3,script)        
        while (re.search(pattern1,script)!=None): #capture overlapping instances of the pattern
            script=re.sub(pattern1,self.rep1,script)
        while (re.search(pattern0,script)!=None): #capture overlapping instances of the pattern
            script=re.sub(pattern0,self.rep0,script)
            
        if "using cells as value" in script:
            script=re.sub(pattern2,self.rep2,script)
            script=script.replace("using cells as value","\n")
        script=re.sub(pattern4,self.rep4,script)
        script=re.sub(pattern5,self.rep5,script)
        script=re.sub(pattern6,self.rep6,script)
        script=re.sub(pattern7,self.rep7,script)
#        a='import openpyxl as xl\nwb=xl.load_workbook("'+self.WBNames[0]+'")\n'
        a='wb=self.scr.edit.activeWS\n'
        script=a+script+"\nself.locals=locals()"
#        script=re.sub('\n','\n\t',script)
        return script
        
    def rep7(self,a):
        a=a.group()
#        print a
        z=re.findall(r'\s',a)
        b=re.search(r'\d+',a)
        c=z[0]+"self.scr.edit.WSindex="+b.group()+z[1]
 #       print c
        return c
    def rep6(self,a):
        a=a.group()
        z=re.findall(r'\s',a)
        b=re.search(r'[\'"].+[\'"]',a)
#        print b
        c="wb=self.addWS("+b.group()+")"
#        print c
        return z[0]+c+z[1]
#    def rep5(self,a):
#        a=a.group()
#        a=re.sub(r'd+',self.rep5_1,a)
#        b=re.search(r'\$.*\$',a)
#        c=re.findall(r'\s',a)
#        b=b.group()[1:-1]
#        a=c[0]+b+c[1]
#        return a
#    def rep5_1(self,a):
#        a=a.group()
#        b="$self.scr.edit.WBindex="+a+"$"
#        return b

    def rep5(self,a):
        a=a.group()
#        print a
        z=re.findall(r'\s',a)
        b=re.search(r'\d+',a)
        c=z[0]+"self.scr.edit.WBindex="+b.group()+z[1]  #note that this DOES NOT change the active worksheet, possibly leading to an out of bounds error if my user is not careful
 #       print c
        return c
        
    def rep4(self,a):
        a=a.group()
        z=re.findall(r'\s',a)
        b=re.search(r'[\'"].+[\'"]',a)
#        print b
        c="wb=self.addWB("+b.group()+")"
#        a=re.sub(r'".+"',self.rep4_1,a)
        return z[0]+c+z[1]
#    def rep4_1(self,a):
#        return "wb=self.addWB("+a+")"
#        a=a.group()
#        self.scr.edit.activeWB.append(xl.Workbook())
#        d=self.scr.edit.activeWB[len(self.widget1.scr.edit.activeWB)-1].worksheets
##        d=self.widget1.scr.edit.activeWB.worksheets
#        self.scr.edit.activeWS.append(d)
##        print type(self.widget1.scr.edit.activeWS[0])
##        self.widget1.scr.edit.activeWS=d
#        z=QTableWidget(0,0,self)
#        v=QTabWidget()
#        v.currentChanged.connect(self.switch_tabs_ws)
#        s=self.scr.edit.activeWS[0][0]
#        for i in s['A1':'Z100']:
#                for j in i:
#                    j.value=''
#        v.addTab(z,"NewSheet")
#        self.table.addTab(v,a)
#        self.updateUI()
#        return ""
#        
    def rep0(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'[A-Z]+\d+',self.rep1_0,a)
        #need to implement an easy user command to change active ws
        return a
    def rep1_0(self,matchobj1):
        a=matchobj1.group()
        rep='["'+a+'"]'
        return rep
        
    def rep1(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'[A-Z]+\d+',self.rep1_1,a)
        #need to implement an easy user command to change active ws
        return a
    def rep1_1(self,matchobj1):
        a=matchobj1.group()
        rep='wb[self.scr.edit.WBindex][self.scr.edit.WSindex]["'+a+'"]'
        return rep
    def rep2(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'wb\[self.scr.edit.WBindex\]\[self.scr.edit.WSindex\]\["[A-Z]+\d+"\]',self.rep2_2,a)
        #need to implement an easy user command to change active ws
        return a
    def rep2_2(self,matchobj1):
        a=matchobj1.group()
        rep=a+'.value'
        return rep
    def rep3(self,b):
        a=b.group()
        a=a.replace(' ','')
        c=re.findall(r'\d+',a)
        d=re.findall(r'[A-Z]+',a)
        for i in range(0,len(c)):
            c[i]=int(c[i])
        c.sort()
        d.sort()
        return encodeindex.decodeList(c[0],c[1],d[0],d[1])
        
#    def highlight(self):  #http://stackoverflow.com/questions/13981824/how-can-i-find-a-substring-and-highlight-it-in-qtextedit
#        cursor = self.textCursor()
#        # Setup the desired format for matches
#        format = QTextCharFormat()
#        format.setBackground(QBrush(QColor("red")))
#        # Setup the regex engine
#        pattern = r'[ \t\n\r\f\v+-=*()][A-Z]+\d+[ \t\n\r\f\v+-=*()\.]'
#        regex = QRegExp(pattern)
#        # Process the displayed document
#        pos = 0
#        index = regex.indexIn(self.document().toPlainText(), pos)
#        print index
#        while (index != -1):
#            # Select the matched text and apply the desired format
#            cursor.setPosition(index)
#            cursor.movePosition(QTextCursor.EndOfWord, 1)
#            cursor.mergeCharFormat(format)
#            # Move to the next match
#            pos = index + regex.matchedLength()
#            index = regex.indexIn(self.toPlainText(), pos)

        
    def getInput(self):
        #method for parsing and colorizing keywords for cell names
        script=self.document().toHtml()
        script='\n'+script+'\n'
#        print script
        pattern1=r'[ \t\n\r\f\1v+-=*()<>][A-Z]+\d+[ \t\n\r\f\v+-=*()\.><]'
#        pattern2=r'=(.*wb.active\["[A-Z]+\d+"\].*)+'
        pattern3=r'\[\s*[A-Z]+\d+\s*-\s*[A-Z]+\d+\s*\]'
        re.MULTILINE=True
        script=re.sub(pattern3,self.crep3,script)        
#        while (re.search(pattern1,script)!=None): #capture overlapping instances of the pattern
        script=re.sub(pattern1,self.crep1,script)
#        script=re.sub(pattern2,self.crep2,script)
        valstate="using cells as value"
        script=script.replace(valstate,'<span style="color:blue;">'+valstate+'</span>')
        script=script.replace("np.",'<span style="color:blue;">'+"np."+'</span>')
        script=script.replace("xl.",'<span style="color:blue;">'+"xl."+'</span>')
        script=script.replace("setActiveWorksheet",'<span style="color:blue;">'+"setActiveWorksheet"+'</span>')
        script=script.replace("setActiveWorbook",'<span style="color:blue;">'+"setActiveWorkbook"+'</span>')
        script=script.replace("newWorksheet",'<span style="color:blue;">'+"newWorksheet"+'</span>')
        script=script.replace("newWorkbook",'<span style="color:blue;">'+"newWorkbook"+'</span>')

        return script
        
        

    def crep1(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'[A-Z]+\d+',self.crep1_1,a)
        #need to implement an easy user command to change active ws
        return a
    def crep1_1(self,matchobj1):
        a=matchobj1.group()
        a='<span style="color:blue;">'+a+'</span>'
        rep=a
        return rep
        
#    def crep2(self,matchobj):
#        a=matchobj.group()
#        a=re.sub(r'wb.active\["[A-Z]+\d+"\]',self.crep2_2,a)
#        #need to implement an easy user command to change active ws
        return a
    def crep2_2(self,matchobj1):
        a=matchobj1.group()
        rep=a+'.value'
        return rep
    def crep3(self,b):
        a=b.group()
        a='<span style="color:blue;">'+a+'</span>'
        return a

"""NOT MY CODE"""
#code from https://john.nachtimwald.com/2009/08/15/qtextedit-with-line-numbers/
#MIT license


#I updated this to work with pyqt5 and added my own subclasses into it

class LineTextWidget(QFrame):
 
    class NumberBar(QWidget):
 
        def __init__(self, *args):
            QWidget.__init__(self, *args)
            self.edit = None
            # This is used to update the width of the control.
            # It is the highest line that is currently visibile.
            self.highest_line = 0
 
        def setTextEdit(self, edit):
            self.edit = edit
 
        def update(self, *args):
            '''
            Updates the number bar to display the current set of numbers.
            Also, adjusts the width of the number bar if necessary.
            '''
            # The + 4 is used to compensate for the current line being bold.
            width = self.fontMetrics().width(str(self.highest_line)) + 4
            if self.width() != width:
                self.setFixedWidth(width)
            QWidget.update(self, *args)
 
        def paintEvent(self, event):
            contents_y = self.edit.verticalScrollBar().value()
            page_bottom = contents_y + self.edit.viewport().height()
            font_metrics = self.fontMetrics()
            current_block = self.edit.document().findBlock(self.edit.textCursor().position())
 
            painter = QPainter(self)
 
            line_count = 0
            # Iterate over all text blocks in the document.
            block = self.edit.document().begin()
            while block.isValid():
                line_count += 1
 
                # The top left position of the block in the document
                position = self.edit.document().documentLayout().blockBoundingRect(block).topLeft()
 
                # Check if the position of the block is out side of the visible
                # area.
                if position.y() > page_bottom:
                    break
 
                # We want the line number for the selected line to be bold.
                bold = False
                if block == current_block:
                    bold = True
                    font = painter.font()
                    font.setBold(True)
                    painter.setFont(font)
 
                # Draw the line number right justified at the y position of the
                # line. 3 is a magic padding number. drawText(x, y, text).
                painter.drawText(self.width() - font_metrics.width(str(line_count)) - 3, round(position.y()) - contents_y + font_metrics.ascent(), str(line_count))
 
                # Remove the bold style if it was set previously.
                if bold:
                    font = painter.font()
                    font.setBold(False)
                    painter.setFont(font)
 
                block = block.next()
 
            self.highest_line = line_count
            painter.end()
 
            QWidget.paintEvent(self, event)
 
 
    def __init__(self, *args):
        QFrame.__init__(self, *args)
 
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
 
        self.edit = script_box()
        self.edit.setFrameStyle(QFrame.NoFrame)
        
 
        self.number_bar = self.NumberBar()
        self.number_bar.setTextEdit(self.edit)
 
        hbox = QHBoxLayout(self)
        hbox.setSpacing(0)
#        hbox.setMargin(0)
        hbox.addWidget(self.number_bar)
        hbox.addWidget(self.edit)
 
        self.edit.installEventFilter(self)
        self.edit.viewport().installEventFilter(self)
 
    def eventFilter(self, object, event):
        # Update the line numbers for all events on the text edit and the viewport.
        # This is easier than connecting all necessary singals.
        if object in (self.edit, self.edit.viewport()):
            self.number_bar.update()
            return False
        return QFrame.eventFilter(object, event)
 
    def getTextEdit(self):
        return self.edit
"""MY CODE AGAIN"""


class outputbox(QTextEdit):
    
    def __init__(self, parent=None) :
        super(outputbox, self).__init__(parent) 
        self.setReadOnly(True)
        self.textChanged.connect(self.restrict_len)
        
    def restrict_len(self):
        if self.document().lineCount()>200:
            self.blockSignals(True) #block signals to prevent infinite loops
            a=self.toPlainText()
            a="\n".join(a.split("\n")[(self.document().lineCount()-190):])
            self.setPlainText(a)
            self.blockSignals(False)


class commandLine(QLineEdit):
    
    def __init__(self, parent=None) :
        super(commandLine, self).__init__(parent) 
#        self.setAcceptRichText(False)
        self.setFrame=False
        self.pastcommands=[]
        self.pastcommandsindex=-1



    def keyPressEvent(self, event):
        key = event.key()
        if key == Qt.Key_Up:
            self.pastcommandsindex+=1
#            print self.pastcommandsindex
            if self.pastcommandsindex >= len(self.pastcommands):
                pass
            else:
                self.setText(self.pastcommands[self.pastcommandsindex])
                
        if key == Qt.Key_Down:
            if self.pastcommandsindex <= 0:
                pass
            else:
                self.pastcommandsindex-=1
                self.setText(self.pastcommands[self.pastcommandsindex])
        super(commandLine,self).keyPressEvent(event)#wow I'm surprised this worked

    def getCode(self):
        #method for parsing and returning user input code
        script=self.text()
        
        self.pastcommands.insert(0,script)
        if len(self.pastcommands)>20:
            del self.pastcommands[-1]
        self.pastcommandsindex=-1
        print self.pastcommands
        
        

        script='\n'+script+'\n'
        

        re.MULTILINE=True


        ###########################################WORKS
#    def rep1(self,matchobj):
#        a=matchobj.group()
#        a=re.sub(r'[A-Z]+\d+',self.rep1_1,a)
#        #need to implement an easy user command to change active ws
#        return a
#    def rep1_1(self,matchobj1):
#        a=matchobj1.group()
#        rep='wb[self.scr.edit.WBindex][self.scr.edit.WSindex]["'+a+'"]'
#        return rep
#    def rep2(self,matchobj):
#        a=matchobj.group()
#        a=re.sub(r'wb\[self.scr.edit.WBindex\]\[self.scr.edit.WSindex\]\["[A-Z]+\d+"\]',self.rep2_2,a)
#        #need to implement an easy user command to change active ws
#        return a
#    def rep2_2(self,matchobj1):
#        a=matchobj1.group()
#        rep=a+'.value'
#        return rep
#    def rep3(self,b):
#        a=b.group()
#        a=a.replace(' ','')
#        c=re.findall(r'\d+',a)
#        d=re.findall(r'[A-Z]+',a)
#        for i in range(0,len(c)):
#            c[i]=int(c[i])
#        c.sort()
#        d.sort()
#        return encodeindex.decodeList(c[0],c[1],d[0],d[1])
#        ##############################################Works


        pattern0=r'[ \t\n\r\f\v+-=*()\[\]][A-Z]+\d+[ \t\n\r\f\v+-=*()\.]'
        pattern1=r'[ \t\n\r\f\v+-=*()\[][A-Z]+\d+[ \t\n\r\f\v+-=*()\.]'
        pattern2=r'=(.*wb\[self.scr.edit.WBindex\]\[self.scr.edit.WSindex\]\["[A-Z]+\d+"\].*)+'
        pattern3=r'\[\s*[A-Z]+\d+\s*-\s*[A-Z]+\d+\s*\]'
        pattern4=r'\snewWorkbook\(["\'].+["\']\)\s'
        pattern5=r'\ssetActiveWorkbook\(\d+\)\s'
        pattern6=r'\snewWorksheet\(["\'].+["\']\)\s'
        pattern7=r'\ssetActiveWorksheet\(\d+\)\s'
        re.MULTILINE=True
        script=re.sub(pattern3,self.rep3,script)        
        while (re.search(pattern1,script)!=None): #capture overlapping instances of the pattern
            script=re.sub(pattern1,self.rep1,script)
        while (re.search(pattern0,script)!=None): #capture overlapping instances of the pattern
            script=re.sub(pattern0,self.rep0,script)
            
        if "using cells as value" in script:
            script=re.sub(pattern2,self.rep2,script)
            script=script.replace("using cells as value","\n")
        script=re.sub(pattern4,self.rep4,script)
        script=re.sub(pattern5,self.rep5,script)
        script=re.sub(pattern6,self.rep6,script)
        script=re.sub(pattern7,self.rep7,script)
        
        
        if "=" not in script and "print" not in script:
            script="print "+script
#        a='import openpyxl as xl\nwb=xl.load_workbook("'+self.WBNames[0]+'")\n'
        a='wb=self.scr.edit.activeWS\n'
        script=a+script+"\nself.locals=locals()"
#        script=re.sub('\n','\n\t',script)
        self.setText("")
        return script
        
    def rep7(self,a):
        a=a.group()
#        print a
        z=re.findall(r'\s',a)
        b=re.search(r'\d+',a)
        c=z[0]+"self.scr.edit.WSindex="+b.group()+z[1]
 #       print c
        return c
    def rep6(self,a):
        a=a.group()
        z=re.findall(r'\s',a)
        b=re.search(r'[\'"].+[\'"]',a)
#        print b
        c="wb=self.addWS("+b.group()+")"
#        print c
        return z[0]+c+z[1]

    def rep5(self,a):
        a=a.group()
#        print a
        z=re.findall(r'\s',a)
        b=re.search(r'\d+',a)
        c=z[0]+"self.scr.edit.WBindex="+b.group()+z[1]  #note that this DOES NOT change the active worksheet, possibly leading to an out of bounds error if my user is not careful
 #       print c
        return c
        
    def rep4(self,a):
        a=a.group()
        z=re.findall(r'\s',a)
        b=re.search(r'[\'"].+[\'"]',a)
#        print b
        c="wb=self.addWB("+b.group()+")"
#        a=re.sub(r'".+"',self.rep4_1,a)
        return z[0]+c+z[1]
    def rep0(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'[A-Z]+\d+',self.rep1_0,a)
        #need to implement an easy user command to change active ws
        return a
    def rep1_0(self,matchobj1):
        a=matchobj1.group()
        rep='["'+a+'"]'
        return rep
        
    def rep1(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'[A-Z]+\d+',self.rep1_1,a)
        #need to implement an easy user command to change active ws
        return a
    def rep1_1(self,matchobj1):
        a=matchobj1.group()
        rep='wb[self.scr.edit.WBindex][self.scr.edit.WSindex]["'+a+'"]'
        return rep
    def rep2(self,matchobj):
        a=matchobj.group()
        a=re.sub(r'wb\[self.scr.edit.WBindex\]\[self.scr.edit.WSindex\]\["[A-Z]+\d+"\]',self.rep2_2,a)
        #need to implement an easy user command to change active ws
        return a
    def rep2_2(self,matchobj1):
        a=matchobj1.group()
        rep=a+'.value'
        return rep
    def rep3(self,b):
        a=b.group()
        a=a.replace(' ','')
        c=re.findall(r'\d+',a)
        d=re.findall(r'[A-Z]+',a)
        for i in range(0,len(c)):
            c[i]=int(c[i])
        c.sort()
        d.sort()
        return encodeindex.decodeList(c[0],c[1],d[0],d[1])
#    def rep1(self,matchobj):
#        a=matchobj.group()
#        a=re.sub(r'[A-Z]+\d+',self.rep1_1,a)
#        #need to implement an easy user command to change active ws
#        return a
#    def rep1_1(self,matchobj1):
#        a=matchobj1.group()
#        rep='wb.active["'+a+'"]'
#        return rep
#    def rep2(self,matchobj):
#        a=matchobj.group()
#        a=re.sub(r'wb.active\["[A-Z]+\d+"\]',self.rep2_2,a)
#        #need to implement an easy user command to change active ws
#        return a
#    def rep2_2(self,matchobj1):
#        a=matchobj1.group()
#        rep=a+'.value'
#        return rep
#    def rep3(self,b):
#        a=b.group()
#        a=a.replace(' ','')
#        c=re.findall(r'\d+',a)
#        d=re.findall(r'[A-Z]+',a)
#        for i in range(0,len(c)):
#            c[i]=int(c[i])
#        c.sort()
#        d.sort()
#        return encodeindex.decodeList(c[0],c[1],d[0],d[1])
        
        
#if __name__ == "__main__" :
app = QApplication(sys.argv)
main = MainWindow()
main.show()
app.exec_()
