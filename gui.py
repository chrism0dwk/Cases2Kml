#!/usr/bin/python
# -*- coding: utf-8 -*-

import wx
import csv
from traceback import print_exc
from string import join
from zipfile import ZipFile,ZIP_DEFLATED
from cases2kml import Cases2kml, DateError
from math import floor


class ErrorDialog(wx.MessageDialog):
    def __init__(self,parent,message):
        wx.MessageDialog.__init__(self,parent,message,caption="Error",style=wx.OK | wx.ICON_ERROR)
        self.ShowModal()
        
class InfoDialog(wx.MessageDialog):
    def __init__(self,parent,message):
        wx.MessageDialog.__init__(self,parent,message,caption="Information",style=wx.OK | wx.ICON_INFORMATION)
        self.ShowModal()


class IptFilePanel(wx.Panel):
    def __init__(self,*args,**kwargs):
        wx.Panel.__init__(self,*args,**kwargs)
        
        # Box sizer created first to preserve Z order
        staticBox = wx.StaticBox(self,label="Input file")
        staticBoxSizer = wx.StaticBoxSizer(staticBox,wx.VERTICAL) 
        

        # CSV Input File
        iptFieldLabel = wx.StaticText(self,label="Choose CSV file:",name="iptFieldLabel")
        self.chooseFile = wx.FilePickerCtrl(self,name="chooseFile")

        iptSizer = wx.BoxSizer(wx.HORIZONTAL)    
        iptSizer.Add(item=iptFieldLabel,proportion=0,flag=wx.EXPAND | wx.LEFT, border=10)
        iptSizer.Add(item=self.chooseFile,proportion=1,flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=10)
        
        # Delimiter character
        delimLabel = wx.StaticText(self,label="Delimiter:",name="delimLabel")
        self.delimChar = wx.TextCtrl(self,name="delimChar")
        self.delimChar.SetValue(",")

        # Quotes
        self.isQuoted = wx.CheckBox(self,label="Quote strings:",name="isQuoted",style=wx.CHK_2STATE)
        self.isQuoted.SetValue(True)
        
        self.quoteChar = wx.TextCtrl(self,name="quoteChar")
        self.quoteChar.SetValue("\"")
        self.quoteChar.Disable()

        delimSizer = wx.BoxSizer(wx.HORIZONTAL)
        delimSizer.Add(item=delimLabel,proportion=0,flag=wx.LEFT,border=10)
        delimSizer.Add(item=self.delimChar,proportion=0,flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=10)
        delimSizer.Add(item=self.isQuoted,proportion=0,flag=wx.EXPAND | wx.LEFT, border=10)
        delimSizer.Add(item=self.quoteChar,proportion=0,flag=wx.EXPAND | wx.LEFT, border=10)

        # Main Sizer       
        staticBoxSizer.Add(item=iptSizer,proportion=0,flag=wx.EXPAND | wx.ALL, border=10)
        staticBoxSizer.Add(item=delimSizer,proportion=0,flag=wx.EXPAND | wx.ALL, border=10)  

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(staticBoxSizer,1,flag=wx.EXPAND | wx.ALL, border = 10)
        
        self.SetSizer(mainSizer)
        self.SetAutoLayout(1)
        staticBoxSizer.Fit(self)
        
        # Wiring
        self.Bind(wx.EVT_CHECKBOX,self.OnClickIsQuoted,self.isQuoted)
        
        self.Show()
        
    def OnClickIsQuoted(self,e):
        if self.quoteChar.IsEnabled():
            self.quoteChar.Disable()
        else:
            self.quoteChar.Enable()
            
    def GetDelimChar(self):
        return self.delimChar.GetValue()
    
    def GetIsQuoted(self):
        print "IsQuoted: " + str(self.isQuoted.IsEnabled())
        return self.isQuoted.IsEnabled()
    
    def GetQuoteChar(self):
        return self.quoteChar.GetValue()


class ChooseFieldsPanel(wx.Panel):
    def __init__(self,*args,**kwargs):
        wx.Panel.__init__(self,*args,**kwargs)
        
        coordBox = wx.StaticBox(self,label="Data fields",name="coordBox")
        xCoordLabel = wx.StaticText(self,label="X Coordinates")
        yCoordLabel = wx.StaticText(self,label="Y Coordinates")
        dateLabel = wx.StaticText(self,label="Date field")
        
        self.xCoordField = wx.ListBox(self,name="xCoord",style=wx.LB_SINGLE | wx.LB_NEEDED_SB)
        self.yCoordField = wx.ListBox(self,name="yCoord",style=wx.LB_SINGLE | wx.LB_NEEDED_SB)
        self.dateField = wx.ListBox(self,name="dateField",style=wx.LB_SINGLE | wx.LB_NEEDED_SB)
        
        # Date formatting
        dateBox = wx.StaticBox(self,label="Date settings", name="dateBox")
        dateFmtLabel = wx.StaticText(self,label="Date field format:",name="dateFmtLabel")
        self.dateFmt = wx.TextCtrl(self,value="%Y-%m-%d",name="dateFmt")
        dateFmtHelp = wx.Button(self,wx.ID_HELP,label="Help",name="dateFmtHelp")
        aggrUnitLabel = wx.StaticText(self,label="Aggregation unit:",name="aggrUnitLabel")
        self.aggrUnitD = wx.RadioButton(self,label="Day",name="aggrUnitM",style=wx.RB_GROUP)
        self.aggrUnitM = wx.RadioButton(self,label="Month",name="aggrUnitM")
        self.aggrUnitY = wx.RadioButton(self,label="Year",name="aggrUnitY")
        self.aggrUnitY.SetValue(True)
        aggrCountLabel = wx.StaticText(self,label="Num time units in aggregation:")
        self.aggrCount = wx.SpinCtrl(self,value="1",name="aggrCount",style=wx.SP_ARROW_KEYS,min=1, max=10000)
        
        # Sizers
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        fieldSizer = wx.FlexGridSizer(2,3,5,10)        
        
        fieldSizer.AddGrowableCol(0,1)
        fieldSizer.AddGrowableCol(1,1)
        fieldSizer.AddGrowableCol(2,1)
        
        fieldSizer.AddGrowableRow(0,1)
        fieldSizer.AddGrowableRow(1,1)
        
        fieldSizer.Add(xCoordLabel)
        fieldSizer.Add(yCoordLabel)
        fieldSizer.Add(dateLabel)
        fieldSizer.Add(self.xCoordField,1,flag=wx.EXPAND | wx.ALL, border=5)
        fieldSizer.Add(self.yCoordField,1,flag=wx.EXPAND | wx.ALL, border=5)
        fieldSizer.Add(self.dateField,1,flag=wx.EXPAND | wx.ALL, border=5)
        
        coordSizer = wx.StaticBoxSizer(coordBox,wx.VERTICAL)
        coordSizer.Add(fieldSizer,1,wx.EXPAND)
        
        # Date format
        dateFmtSizer = wx.BoxSizer(wx.HORIZONTAL)
        dateFmtSizer.Add(dateFmtLabel)
        dateFmtSizer.Add(self.dateFmt,flag=wx.LEFT,border=10)
        dateFmtSizer.Add(dateFmtHelp,flag=wx.LEFT,border=10)
        
        # Aggregation units
        aggrUnitSizer = wx.BoxSizer(wx.HORIZONTAL)
        aggrUnitSizer.Add(aggrUnitLabel)
        aggrUnitSizer.Add(self.aggrUnitD,0,flag=wx.LEFT,border=10)
        aggrUnitSizer.Add(self.aggrUnitM,0,flag=wx.LEFT,border=10)
        aggrUnitSizer.Add(self.aggrUnitY,0,flag=wx.LEFT,border=10)
        
        aggrCountSizer = wx.BoxSizer(wx.HORIZONTAL)
        aggrCountSizer.Add(aggrCountLabel)
        aggrCountSizer.Add(self.aggrCount,0,flag=wx.LEFT,border=10)

        # Layout for date stuff        
        dateSizer = wx.StaticBoxSizer(dateBox,wx.VERTICAL)
        dateSizer.Add(dateFmtSizer,flag = wx.EXPAND | wx.ALL,border=10)
        dateSizer.Add(aggrUnitSizer,flag=wx.EXPAND | wx.ALL, border=10)
        dateSizer.Add(aggrCountSizer,flag = wx.EXPAND | wx.ALL, border=10)
        
        
        mainSizer.Add(coordSizer,0,wx.EXPAND | wx.ALL, border=10)
        mainSizer.Add(dateSizer,1,wx.EXPAND | wx.ALL, border=10)
        
        self.SetSizer(mainSizer)
        self.SetAutoLayout(1)
        mainSizer.Fit(self)
        
        # Wiring
        self.Bind(wx.EVT_BUTTON, self.OnDateHelp, dateFmtHelp)
        
        self.Show()
        
    def OnDateHelp(self,e):
        
        message = """The date format used by Cases2Kml follows the standard format for strftime().  Some useful format codes are:\n\n
%d - gives the day of the month [1,31]\n\n
%m - gives the month number [1,12]\n\n
%Y - gives the year in full (eg. 1999, 2009)\n\n
%y - gives the year without century (eg. 99, 09)\n\n
For more information, see documentation for strftime()"""

        helpWindow = wx.MessageDialog(self,message=message,caption="Help on date format",style=wx.CANCEL)
        helpWindow.ShowModal()
        
    def getAggrUnit(self):
        if self.aggrUnitD.GetValue() == True:
            return 'D'
        elif self.aggrUnitM.GetValue() == True:
            return 'M'
        else:
            return 'Y'
        
    def getAggrCount(self):
        return int(self.aggrCount.GetValue())
        
    def populateListBoxes(self,items):
        self.xCoordField.Set(items)
        self.yCoordField.Set(items)
        self.dateField.Set(items)


class ChooseOutputOptions(wx.Panel):
    def __init__(self,*args,**kwargs):
        
        self.mag = 1.0 # Used to set point magnification
        
        wx.Panel.__init__(self,*args,**kwargs)
        
        staticBox = wx.StaticBox(self,label="Display options", name="staticBox")
        staticBoxSizer = wx.StaticBoxSizer(staticBox,wx.VERTICAL)
       
        magLabel = wx.StaticText(self,label="Map point magnification: ",name="magLabel")
        self.magCtrl = wx.TextCtrl(self,value="1.0",size=(40,-1),name="mag")
        self.magSpinner = wx.SpinButton(self,name="magSpin",style=wx.SP_ARROW_KEYS)
        self.magSpinner.SetRange(0,1000)
        self.magSpinner.SetValue(10)
        
        colLabel = wx.StaticText(self,label="Map point colour: ",name="colLabel")
        self.col = wx.ColourPickerCtrl(self,col=wx.RED)
        
        # Layout        
        magSizer = wx.BoxSizer(wx.HORIZONTAL)
        magSizer.Add(magLabel)
        magSizer.Add(self.magCtrl)
        magSizer.Add(self.magSpinner,0,wx.LEFT,border=8)
        magSizer.Add(colLabel,0,wx.LEFT, border=10)
        magSizer.Add(self.col,0)

        staticBoxSizer.Add(magSizer)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(staticBoxSizer,0,wx.EXPAND | wx.ALL, border=10)
        
        self.Layout()
        self.SetSizer(mainSizer)
        self.SetAutoLayout(1)
        mainSizer.Fit(self)
        
        # Wiring
        self.Bind(wx.EVT_SPIN_UP,self.OnMagUp,self.magSpinner)
        self.Bind(wx.EVT_SPIN_DOWN, self.OnMagDown,self.magSpinner)
        self.Bind(wx.EVT_TEXT, self.OnMagCtrlUpdate,self.magCtrl)

        self.Show()
        
    def OnMagUp(self,e):
        mag = self.mag + 0.1
        self.magCtrl.SetValue(str(mag))
        self.magCtrl.Refresh()
        
    def OnMagDown(self,e):
        if self.mag >= 0.1:
            mag = self.mag - 0.1
            self.magCtrl.SetValue(str(mag))
            self.magCtrl.Refresh()
        
    def OnMagCtrlUpdate(self,e):
        if self.magCtrl.GetValue() == '':
            return
        try:
            mag = float(self.magCtrl.GetValue()) * 10
            mag = floor(mag) / 10
            if mag < 0 or mag > 1000:
                raise ValueError
            self.mag = mag
            self.magSpinner.SetValue(mag*10)
        except ValueError as err:
            ErrorDialog(self,"Value error! Magnification must be a number >= 0")
            
        
    def getMagnification(self):
        return float( self.mag )
    
    def getKMLColour(self):
        colour = self.col.GetColour()
        
        alpha = colour.Alpha()
        blue = colour.Blue()
        green = colour.Green()
        red = colour.Red()
        
        return "%02X%02X%02X%02X" % (alpha,blue,green,red)
        
    
    def getOutputFileName(self):
        return str( self.optFilePicker.GetPath() )
    


        
class ButtonPanel(wx.Panel):
    def __init__(self,*args,**kwargs):
        wx.Panel.__init__(self,*args,**kwargs)
        
        buttonSizer = wx.StdDialogButtonSizer()

        self.saveButton = wx.Button(self,wx.ID_SAVE,label="Save")
        self.quitButton = wx.Button(self,wx.ID_EXIT,label="Quit")
        
        buttonSizer.AddButton(self.saveButton)
        buttonSizer.SetCancelButton(self.quitButton)
        buttonSizer.Realize()

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(buttonSizer,1,flag=wx.EXPAND | wx.ALL, border = 10)
        self.SetSizerAndFit(mainSizer)
        
        self.Show()

class MainWindow(wx.Frame):
    def __init__(self,parent,title):

        # Housekeeping
        self.fieldMap = {'x': None, 'y': None, 'date': None}

        wx.Frame.__init__(self,parent,title=title)
        
        self.CreateStatusBar()

        filemenu = wx.Menu()
        menuAbout = filemenu.Append(wx.ID_ABOUT, "&About"," Information about this program")
        filemenu.AppendSeparator()
        menuExit = filemenu.Append(wx.ID_EXIT,"E&xit"," Quit")

        menuBar = wx.MenuBar()
        menuBar.Append(filemenu, "&File")
        self.SetMenuBar(menuBar)
        
        # Main Panels
        panelStyle = wx.TAB_TRAVERSAL | wx.BORDER_NONE
        self.iptFilePanel = IptFilePanel(self,name="iptFilePanel",style=panelStyle)
        self.chooseFieldPanel = ChooseFieldsPanel(self,name="chooseFieldPanel",style=panelStyle)
        self.chooseOutputOptionsPanel = ChooseOutputOptions(self,name="chooseOutputOptions",style=panelStyle)
        self.buttonPanel = ButtonPanel(parent=self,style=panelStyle,name="buttonPanel")

        # Wiring
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        self.Bind(wx.EVT_BUTTON, self.OnSave, self.buttonPanel.saveButton)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.buttonPanel.quitButton)
        self.Bind(wx.EVT_FILEPICKER_CHANGED, self.OnChooseIptFile, self.iptFilePanel.chooseFile)
        self.Bind(wx.EVT_SIZE, self.OnResize)

        # Layout
        self.mainSizer = wx.BoxSizer(wx.VERTICAL)
        self.mainSizer.Add(item=self.iptFilePanel,proportion=0,flag = wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border = 0)
        self.mainSizer.Add(item=self.chooseFieldPanel,proportion=0,flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border = 0)
        self.mainSizer.Add(item=self.chooseOutputOptionsPanel,proportion=0,flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border = 0)
        self.mainSizer.Add(item=self.buttonPanel,proportion=0,flag = wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border = 0)
        
        self.SetSizer(self.mainSizer)
        self.SetAutoLayout(True)
        self.Fit()
        
        self.SetMinSize(self.GetSize())
        
        self.Show(True)
        
    def OnAbout(self,e):
        message = """Cases2Kml version 1.0-6beta

Author: Chris Jewell <chris.jewell@warwick.ac.uk> (c) 2010

License: GPL 3

Purpose: Converts CSV files of disease cases into time-aggregated Google Earth .kmz files
"""
        dlg = wx.MessageDialog(self,message=message,caption="About Cases2Kml",style=wx.OK)
        dlg.ShowModal()

    def OnExit(self,e):
        self.Close(True)
        
    def OnChooseIptFile(self,e):
            try:
                self.getCSVHeader()
            except IOError:
                print_exc()
                ErrorDialog(self,"File not found! Please check your path and try again.")
                return
            except Exception as inst:
                print_exc()
                ErrorDialog(self,"Unknown error.  Are you sure this is a CSV file?")
                return
            e.Skip()
        
    def OnResize(self,e):
        statusText = "(" + str(self.GetSize().GetWidth()) + "," + str(self.GetSize().GetHeight()) + ")"
        self.GetStatusBar().SetStatusText(statusText)
        e.Skip()
        
    def OnSave(self,e):
        
        # Set field choices
        try:
            self.setFieldMap()
        except Exception as e:
            ErrorDialog(self, message="One or more CSV fields not set!")
            return
        
        # Check date field
        if self.chooseFieldPanel.dateFmt.IsEmpty():
            ErrorDialog(self,message="Date format field empty!")
            return
        
        # Run the save file dialog
        saveDlg = wx.FileDialog(self,message="Save file",wildcard="Google Earth files (*.kmz;*.kml)|*.kmz;*.kml",style=wx.FD_SAVE)
        saveDlg.ShowModal()
        
        if saveDlg.GetPath() != "":
            self.generateKML(saveDlg.GetPath())
            

    def getCSVHeader(self):
        filename = str( self.iptFilePanel.chooseFile.GetPath() )
        delimiter = str( self.iptFilePanel.delimChar.GetValue() )
        useQuotes = bool( self.iptFilePanel.isQuoted.GetValue() )
        quoteChar = str( self.iptFilePanel.quoteChar.GetValue() )
        
        iptFile = open(filename, "r")
        myReader = csv.DictReader(iptFile, delimiter=delimiter, quoting=useQuotes, quotechar=quoteChar)
        self.chooseFieldPanel.populateListBoxes(myReader.fieldnames)
        iptFile.close()
        
    def setFieldMap(self):        
        self.fieldMap['x'] = self.chooseFieldPanel.xCoordField.GetSelection()
        self.fieldMap['y'] = self.chooseFieldPanel.yCoordField.GetSelection()
        self.fieldMap['date'] = self.chooseFieldPanel.dateField.GetSelection()
        
        if wx.NOT_FOUND in self.fieldMap.values():
            raise ValueError("One or more fields not set")
            

    def generateKML(self,outputFileName):
        
        progress = wx.ProgressDialog("Processing","Reading CSV. Please wait...",style=wx.PD_SMOOTH | wx.PD_REMAINING_TIME)
        
        try:
            kmz = ZipFile( outputFileName,"w",ZIP_DEFLATED)
        except IOError as err:
            ErrorDialog(self,"Could not open output file: " + err.args[1])
            progress.Destroy()
            return

        try:
            converter = Cases2kml( self.chooseFieldPanel.getAggrUnit(),self.chooseFieldPanel.getAggrCount(), self.chooseOutputOptionsPanel.getMagnification(), self.chooseOutputOptionsPanel.getKMLColour() )
            converter.readCSV( self.iptFilePanel.chooseFile.GetPath(),
                               self.fieldMap,
                               self.chooseFieldPanel.dateFmt.GetValue(),
                               self.iptFilePanel.GetQuoteChar(),
                               self.iptFilePanel.GetDelimChar(),
                               progress.Pulse )                 
            progress.Pulse("Writing KMZ file.  Please wait...")
            kmz.writestr("doc.kml",converter.serialize( outputFileName, progress.Update ))
            InfoDialog(self,"Conversion complete")
        
        except DateError as err:
            ErrorDialog(self,"Could not parse the date field.  Please check your date format settings.")
            
        except IOError as err:
            print_exc()
            ErrorDialog(self,"Cannot open CSV file.  Please check the file permissions.")
  
        except Exception as e:
            ErrorDialog(self,"An unknown error prevented the .kmz file from being written:\n\n \"" + str(e) + "\"\n\nPlease check your settings.")
            print_exc()
            
        kmz.close()    
        progress.Destroy()        




if __name__ == "__main__":
    app = wx.App(False)
    frame = MainWindow(None,"Cases2KML")
    app.MainLoop()
