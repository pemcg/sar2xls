#!/usr/bin/env python
#
#-------------------------------------------------------------------------------------------------------------
# Author:   srijit@yahoo.com
#           modified by Peter McGowan
#
# Version:	0.1
#           0.2     02-Jul-09   PEMcG   Improved error handling
#           0.3     13-Jul-09   PEMcG   Added comments to plotdata, and debugged save
#           0.4     16-May-11   PEMcG   Added lastcellincolumn function
#
# Revision History
#
#-------------------------------------------------------------------------------------------------------------

import win32com.client.dynamic
import sys

class UseExcel(object):
    """Python Excel Interface. It provides methods for accessing the
    basic functionality of MS Excel 97/2000 from Python.

    This interface uses dynamic dispatch objects. Most necessary constants
    are embedded in the code. For others we need to run makepy.py.
    """

    __slots__ = ("xlapp", "xlbook")

#-------------------------------------------------------------------------------------------------------------
    def __init__(self, fileName=None):
        #
        #e.g. xlFile = useExcel("e:\\python23\myfiles\\testExcel1.xls")
        #
        self.xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
        self.xlapp.SheetsInNewWorkbook = 1
        if fileName:
            #try:
                self.xlbook = self.xlapp.Workbooks.Open(fileName)

        else:
            self.xlbook = self.xlapp.Workbooks.Add()
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def save(self, NewFileName=None):
        if NewFileName:
            self.xlbook.SaveAs(NewFileName)
        else:
            self.xlbook.Save()
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def close(self):
        self.xlbook.Close(SaveChanges=False)
        del self.xlapp
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def show(self):
        self.xlapp.Visible = True
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def hide(self):
        self.xlapp.Visible = False
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def getcell(self, sheet, cellAddress):

        """Get value of one cell.
        Description of parameters (self explanatory parameters are not described):
            sheet       -   name of the excel worksheet
            cellAddress -   tuple of integers (row, cloumn) or string "ColumnRow"
                            e.g. (3,4) or "D3"
        """

        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(cellAddress,str)):
            return sht.Range(cellAddress).Value
        elif (isinstance(cellAddress,tuple)):
            row = cellAddress[0]
            col = cellAddress[1]
            return sht.Cells(row, col).Value
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def lastcellincolumn(self, sheet, column):

        """Find the last used cell in a column
        Description of parameters (self explanatory parameters are not described):
            sheet       -   name of the excel worksheet
            column      -   Column letter to find max row in
        """
        xlUp = -4162
        sht = self.xlbook.Worksheets(sheet)
        LastRow = sht.Range(column + "65536").End(xlUp).Row
        return LastRow

#-------------------------------------------------------------------------------------------------------------
    def setcellvalue(self, sheet, value, cellAddress, fontStyle=("Regular",), fontName="Arial", fontSize=12, fontColor=1):

        """Set value of one cell.
        Description of parameters (self explanatory parameters are not described):
            sheet       -   name of the excel worksheet
            value       -   The cell value. it can be a number, string etc.
            cellAddress -   tuple of integers (row, cloumn) or string "ColumnRow"
                            e.g. (3,4) or "D3"
            fontStyle   -   tuple. Combination of Regular, Bold, Italic, Underline
                            e.g. ("Regular", "Bold", "Italic")
            fontColor   -   ColorIndex. Refer ColorIndex property in Microsoft Excel Visual Basic Reference
        """

        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(cellAddress,str)):
            sht.Range(cellAddress).Value = value
            sht.Range(cellAddress).Font.Size = fontSize
            sht.Range(cellAddress).Font.ColorIndex = fontColor
            for i, item in enumerate(fontStyle):
                if (item.lower() == "bold"):
                    sht.Range(cellAddress).Font.Bold = True
                elif (item.lower() == "italic"):
                    sht.Range(cellAddress).Font.Italic = True
                elif (item.lower() == "underline"):
                    sht.Range(cellAddress).Font.Underline = True
                elif (item.lower() == "regular"):
                    sht.Range(cellAddress).Font.FontStyle = "Regular"
            sht.Range(cellAddress).Font.Name = fontName
        elif (isinstance(cellAddress,tuple)):
            row = cellAddress[0]
            col = cellAddress[1]
            sht.Cells(row, col).Value = value
            sht.Cells(row, col).Font.FontSize = fontSize
            sht.Cells(row, col).Font.ColorIndex = fontColor
            for i, item in enumerate(fontStyle):
                if (item.lower() == "bold"):
                    sht.Range(cellAddress).Font.Bold = True
                elif (item.lower() == "italic"):
                    sht.Range(cellAddress).Font.Italic = True
                elif (item.lower() == "underline"):
                    sht.Range(cellAddress).Font.Underline = True
                elif (item.lower() == "regular"):
                    sht.Range(cellAddress).Font.FontStyle = "Regular"
            sht.Cells(row, col).Font.Name = fontName
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def setcellformula(self, sheet, formula, cellAddress, fontStyle=("Regular",), fontName="Arial", fontSize=12, fontColor=1):

        """Set value of one cell.
        Description of parameters (self explanatory parameters are not described):
            sheet       -   name of the excel worksheet
            formula     -   The cell formula
            cellAddress -   tuple of integers (row, cloumn) or string "ColumnRow"
                            e.g. (3,4) or "D3"
            fontStyle   -   tuple. Combination of Regular, Bold, Italic, Underline
                            e.g. ("Regular", "Bold", "Italic")
            fontColor   -   ColorIndex. Refer ColorIndex property in Microsoft Excel Visual Basic Reference
        """

        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(cellAddress,str)):
            sht.Range(cellAddress).Formula = formula
            sht.Range(cellAddress).Font.Size = fontSize
            sht.Range(cellAddress).Font.ColorIndex = fontColor
            for i, item in enumerate(fontStyle):
                if (item.lower() == "bold"):
                    sht.Range(cellAddress).Font.Bold = True
                elif (item.lower() == "italic"):
                    sht.Range(cellAddress).Font.Italic = True
                elif (item.lower() == "underline"):
                    sht.Range(cellAddress).Font.Underline = True
                elif (item.lower() == "regular"):
                    sht.Range(cellAddress).Font.FontStyle = "Regular"
            sht.Range(cellAddress).Font.Name = fontName
        elif (isinstance(cellAddress,tuple)):
            row = cellAddress[0]
            col = cellAddress[1]
            sht.Cells(row, col).Value = value
            sht.Cells(row, col).Font.FontSize = fontSize
            sht.Cells(row, col).Font.ColorIndex = fontColor
            for i, item in enumerate(fontStyle):
                if (item.lower() == "bold"):
                    sht.Range(cellAddress).Font.Bold = True
                elif (item.lower() == "italic"):
                    sht.Range(cellAddress).Font.Italic = True
                elif (item.lower() == "underline"):
                    sht.Range(cellAddress).Font.Underline = True
                elif (item.lower() == "regular"):
                    sht.Range(cellAddress).Font.FontStyle = "Regular"
            sht.Cells(row, col).Font.Name = fontName
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def getrange(self, sheet, rangeAddress):

        """Returns a tuple of tuples from a range of cells. Each tuple corresponds to a row in excel sheet.

        Description of parameters (self explanatory parameters are not described):
            sheet           -   name of the excel worksheet
            rangeAddress    -   tuple of integers (row1,col1,row2,col2) or "cell1Address:cell2Address"
                                row1,col1 refers to first cell
                                row2,col2 refers to second cell
                                e.g. (1,2,5,7) or "B1:G5"
        """

        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,str)):
            return sht.Range(rangeAddress).Value
        elif (isinstance(rangeAddress,tuple)):
            row1 = rangeAddress[0]
            col1 = rangeAddress[1]
            row2 = rangeAddress[2]
            col2 = rangeAddress[3]
            return sht.Range(sht.Cells(row1, col1), sht.Cells(row2,col2)).Value
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def setrange(self, sheet, topRow, leftCol, data):

        """Sets range of cells with values from data. data is a tuple of tuples.
            Each tuple corresponds to a row in excel sheet.

        Description of parameters (self explanatory parameters are not described):
            sheet   -   name of the excel worksheet
            topRow  -   row number (integer data type)
            leftCol -   column number (integer data type)
        """

        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        sht = self.xlbook.Worksheets(sheet)
        sht.Range(sht.Cells(topRow, leftCol), sht.Cells(bottomRow, rightCol)).Value = data
        return (bottomRow, rightCol)
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def setcellalign(self, sheet, cellAddress, alignment):

        """Aligns the contents of the cell.

        Description of parameters (self explanatory parameters are not described):
            sheet       -   name of the excel worksheet
            cellAddress -   tuple of integers (row, cloumn) or string "ColumnRow"
                            e.g. (3,4) or "D3"
            alignment   -   "Left", "Right" or "center"
        """

        if (alignment.lower() == "left"):
            alignmentValue = 2
        elif ((alignment.lower() == "center") or (alignment.lower() == "centre")):
            alignmentValue = 3
        elif (alignment.lower() == "right"):
            alignmentValue = 4
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(cellAddress,str)):
            sht.Range(cellAddress).HorizontalAlignment = alignmentValue
        elif (isinstance(cellAddress,tuple)):
            row = cellAddress[0]
            col = cellAddress[1]
        sht.Cells(row, col).HorizontalAlignment = alignmentValue
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def addnewworksheetbefore(self, oldSheet, newSheetName):

        """Adds a new excel sheet before the given excel sheet.

        Description of parameters (self explanatory parameters are not described):
            oldSheet        -   Name of the sheet before which a new sheet should be inserted
            newSheetName    -   Name of the new sheet
        """

        sht = self.xlbook.Worksheets(oldSheet)
        self.xlbook.Worksheets.Add(sht).Name = newSheetName
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def deleteworksheet(self, SheetName):

        """Deletes an excel sheet in a workbook.

        Description of parameters (self explanatory parameters are not described):
            SheetName    -   Name of the sheet to delete
        """

        sht = self.xlbook.Worksheets(SheetName)
        sht.Delete()

#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def addnewworksheetafter(self, oldSheet, newSheetName):

        """Adds a new excel sheet after the given excel sheet.

        Description of parameters (self explanatory parameters are not described):
            oldSheet        -   Name of the sheet after which a new sheet should be inserted
            newSheetName    -   Name of the new sheet
        """

        sht = self.xlbook.Worksheets(oldSheet)
        self.xlbook.Worksheets.Add(None,sht).Name = newSheetName
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def insertchart(self, sheet, left, top, width, height):

        """Creates a new embedded chart. Returns a ChartObject object.
            Refer Add Method(ChartObjects Collection) in Microsoft Excel Visual Basic Reference.

        Description of parameters (self explanatory parameters are not described):
            sheet           -   name of the excel worksheet
            left, top       -   The initial coordinates of the new object (in points),
                                relative to the upper-left corner of cell A1 on a worksheet
                                or to the upper-left corner of a chart.
            width, height   -   The initial size of the new object, in points.
                                (point = unit of measurement equal to 1/72 inch.)
        """

        sht = self.xlbook.Worksheets(sheet)
        return sht.ChartObjects().Add(left, top, width, height)
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def plotdata(self, sheet, dataRanges, chartObject, gallery, format=None, plotBy=None,
                        categoryLabels=1, seriesLabels=0, hasLegend=None, title=None,
                        categoryTitle=None, valueTitle=None, extraTitle=None):

        """Plots data using ChartWizard. For details refer ChartWizard method in Microsoft Excel Visual Basic Reference.
            Before using PlotData method InsertChart method should be used.

        Description of parameters:
            sheet       -   name of the excel worksheet. This name should be same as that in InsertChart method
            dataRanges  -   tuple of tuples ((topRow, leftCol, bottomRow, rightCol),). Range of data in excel worksheet to be plotted.
            chartObject -   Embedded chart object returned by InsertChart method.
            chartType   -   Refer plotType variable for available options.
            For remaining parameters refer ChartWizard method in Microsoft Excel Visual Basic Reference:
        """

        sht = self.xlbook.Worksheets(sheet)
        if (len(dataRanges) == 1):
            topRow, leftCol, bottomRow, rightCol = dataRanges[0]
            source = sht.Range(sht.Cells(topRow, leftCol), sht.Cells(bottomRow, rightCol))
        elif (len(dataRanges) > 1):
            topRow, leftCol, bottomRow, rightCol = dataRanges[0]
            source = sht.Range(sht.Cells(topRow, leftCol), sht.Cells(bottomRow, rightCol))
            for count in range(len(dataRanges[1:])):
                topRow, leftCol, bottomRow, rightCol = dataRanges[count+1]
                tempSource = sht.Range(sht.Cells(topRow, leftCol), sht.Cells(bottomRow, rightCol))
                source = self.xlapp.Union(source, tempSource)
        plotType = {
                            "Area" : 1,
                            "Bar" : 2,
                            "Column" : 3,
                            "Line" : 4,
                            "LineMarkers" : 65,
                            "LineMarkersStacked" : 66,
                            "LineStacked" : 63,
                            "Pie" : 5,
                            "Radar" : -4151,
                            "Scatter" : -4169,
                            "Combination" : -4111,
                            "3DArea" : -4098,
                            "3DBar" : -4099,
                            "3DColumn" : -4100,
                            "3DPie" : -4101,
                            "3DSurface" : -4103,
                            "Doughnut" : -4120,
                            "Radar" : -4151,
                            "Bubble" : 15,
                            "Surface" : 83,
                            "Cone" : 3,
                            "3DAreaStacked" : 78,
                            "3DColumnStacked" : 55
                            }
        #gallery = plotType[chartType]
  
        """
            Parameters (from ChartWizard method in Microsoft Excel Visual Basic Reference)
                Source
                    The range that contains the source data for the new chart. If this argument is omitted, Microsoft Office Excel
                    edits the active chart sheet or the selected Chart control on the active worksheet.
                Gallery
                    XlChartType. The chart type.
                Format
                    The option number for the built-in autoformats. Can be a number from 1 through 10, depending on the gallery type.
                    If this argument is omitted, Excel chooses a default value based on the gallery type and data source.
                PlotBy
                    Specifies whether the data for each series is in rows or columns. Can be one of the following XlRowCol constants:
                    xlRows or xlColumns.
                CategoryLabels
                    An integer specifying the number of rows or columns within the source range that contain category labels.
                    Legal values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.
                SeriesLabels
                    An integer specifying the number of rows or columns within the source range that contain series labels.
                    Legal values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.
                HasLegend
                    true to include a legend.
                Title
                    The Chart control title text.
                CategoryTitle
                    The category axis title text.
                ValueTitle
                    The value axis title text
                ExtraTitle
                    The series axis title for 3-D charts or the second value axis title for 2-D charts.
        """
        chartObject.Chart.ChartWizard(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, extraTitle)
        
        nSeries = chartObject.Chart.SeriesCollection().Count
        for chart in range(1,nSeries+1):
            chartObject.Chart.SeriesCollection(chart).Format.Line.Weight = 1

#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def copyrange(self, source, destination):

        """Copy range of data from source range in a sheet to destination range in same sheet or different sheet
            in the same workbook

        Description of parameters (self explanatory parameters are not described):
            source          -   tuple (sheet, rangeAddress)
                                sheet - name of the excel sheet
                                rangeAddress - "cell1Address:cell2Address"
            destination     -   tuple (sheet, destinationCellAddress)
                                destinationCellAddress - string "ColumnRow"
        """

        sourceSht = self.xlbook.Worksheets(source[0])
        destinationSht = self.xlbook.Worksheets(destination[0])
        sourceSht.Range(source[1]).Copy(destinationSht.Range(destination[1]))
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def copyrangetoclipboard(self, source):

        """Copy range of data from source range in a sheet to the clipboard

        Description of parameters (self explanatory parameters are not described):
            source          -   tuple (sheet, rangeAddress)
                                sheet - name of the excel sheet
                                rangeAddress - "cell1Address:cell2Address"
        """

        self.xlbook.Activate()
        sourceSht = self.xlbook.Worksheets(source[0])
        sourceSht.Range(source[1]).Copy()
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def copycolumntoclipboard(self, source):

        """Copy a column of data from source range in a sheet to the clipboard

        Description of parameters (self explanatory parameters are not described):
            source          -   tuple (sheet, rangeAddress)
                                sheet - name of the excel sheet
                                rangeAddress - "cell1Address:cell2Address"
        """

        self.xlbook.Activate()
        sourceSht = self.xlbook.Worksheets(source[0])
        sourceSht.Range(source[1]).EntireColumn.Copy()
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def pasterangefromclipboard(self, destination):

        """Copy range of data from clipboard to destination range in a sheet

        Description of parameters (self explanatory parameters are not described):
            destination     -   tuple (sheet, destinationCellAddress)
                                destinationCellAddress - string "ColumnRow"
        """

        self.xlbook.Activate()
        self.xlbook.Worksheets(destination[0]).Range(destination[1]).Select()
        self.xlbook.Worksheets(destination[0]).Paste()
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def copychart(self, sourceChartObject, destination, delete="N"):

        """Copy chart from source range in a sheet to destination range in same sheet or different sheet

        Description of parameters (self explanatory parameters are not described):
            sourceChartObject   -   Chart object returned by InsertChart method.
            destination         -   tuple (sheet, destinationCellAddress)
                                    sheet - name of the excel worksheet.
                                    destinationCellAddress - string "ColumnRow"
                                    if sheet is omitted and only destinationCellAddress is available
                                    as string data then same sheet is assumed.
            delete              -   "Y" or "N". If "Y" the source chart object is deleted after copy.
                        So if "Y" copy chart is equivalent to move chart.
        """

        if (isinstance(destination,tuple)):
            sourceChartObject.Copy()
            sht = self.xlbook.Worksheets(destination[0])
            sht.Paste(sht.Range(destination[1]))
        else:
            sourceChartObject.Chart.ChartArea.Copy()
            destination.Chart.Paste()
        if (delete.upper() =="Y"):
            sourceChartObject.Delete()
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def hidecolumn(self, sheet, col):

        """Hide a column.

        Description of parameters (self explanatory parameters are not described):
            sheet   -   name of the excel worksheet.
            col     -   column number (integer data)
        """

        sht = self.xlbook.Worksheets(sheet)
        sht.Columns(col).Hidden = True
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def hiderow(self, sheet, row):

        """ Hide a row.

        Description of parameters (self explanatory parameters are not described):
            sheet   -   name of the excel worksheet.
            row     -   row number (integer data)
        """

        sht = self.xlbook.Worksheets(sheet)
        sht.Rows(row).Hidden = True
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def excelfunction(self, sheet, range, function):

        """Access Microsoft Excel worksheet functions. Refer WorksheetFunction Object in Microsoft Excel Visual Basic Reference

        Description of parameters (self explanatory parameters are not described):
            sheet   -   name of the excel worksheet
            range   -   tuple of integers (row1,col1,row2,col2) or "cell1Address:cell2Address"
                        row1,col1 refers to first cell
                        row2,col2 refers to second cell
                        e.g. (1,2,5,7) or "B1:G5"
            For list of functions refer List of Worksheet Functions Available to Visual Basic in Microsoft Excel Visual Basic Reference
        """

        sht = self.xlbook.Worksheets(sheet)
        if isinstance(range,str):
            xlRange = "(sht.Range(" + "'" + range + "'" + "))"
        elif isinstance(range,tuple):
            topRow = range[0]
            leftColumn = range[1]
            bottomRow = range[2]
            rightColumn = range[3]
            xlRange = "(sht.Range(sht.Cells(topRow, leftColumn), sht.Cells(bottomRow, rightColumn)))"
        xlFunction = "self.xlapp.WorksheetFunction." + function + xlRange
        return eval(xlFunction, globals(), locals())
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def clearrange(self, sheet, rangeAddress, contents="Y", format="Y"):

        """Clear the contents of a range of cells.
        Description of parameters (self explanatory parameters are not described):
            sheet           -   name of the excel worksheet
            rangeAddress    -   tuple of integers (row1,col1,row2,col2) or "cell1Address:cell2Address"
                                row1,col1 refers to first cell
                                row2,col2 refers to second cell
                                e.g. (1,2,5,7) or "B1:G5"
            contents        -   "Y" or "N". If "Y" clears the formulas from the range
            format          -   "Y" or "N". If "Y" clears the formatting of the object
        """

        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,str)):
            if (format.upper() == "Y"):
                sht.Range(rangeAddress).ClearFormats()
            if (contents.upper() == "Y"):
                sht.Range(rangeAddress).ClearContents()
        elif (isinstance(rangeAddress,tuple)):
            row1 = rangeAddress[0]
            col1 = rangeAddress[1]
            row2 = rangeAddress[2]
            col2 = rangeAddress[3]
            if (format.upper() == "Y"):
                sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).ClearFormats()
            if (contents.upper() == "Y"):
                sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).ClearContents()
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
    def addcomment(self, sheet, cellAddress, comment=""):

        """Add or delete comment to a cell. If parameter comment is None, delete the comments

        Description of parameters (self explanatory parameters are not described):
            sheet       -   name of the excel worksheet
            cellAddress -   tuple of integers (row, cloumn) or string "ColumnRow"
                            e.g. (3,4) or "D3"
            comment     -   String data. Comment to be added. If None, delete comments
        """

        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(cellAddress,str)):
            if (comment != None):
                sht.Range(cellAddress).AddComment(comment)
            else:
                sht.Range(cellAddress).ClearComments()
        elif (isinstance(cellAddress,tuple)):
            row1 = cellAddress[0]
            col1 = cellAddress[1]
            if (comment != None):
                sht.Cells(row1, col1).AddComment(comment)
            else:
                sht.Cells(row1, col1).ClearComments()
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
def excelapp():
    xlFile = UseExcel("e:\\python23\myfiles\\StudentTabulation.xls")
    xlFile.show()
    xlFile.setcell(sheet="Sheet1", value="Class X Annual Examination",fontName="Arial",
                         cellAddress="D1",fontColor=1, fontStyle=("Bold",), fontSize=16)
    xlFile.setcell(sheet="Sheet1", value="Subject : History", fontName="Arial",
                         cellAddress="D3",fontColor=1)
    data =  (
                ("Sl. No."  ,"Name of Students", "Roll No.", "Marks(out of 100)"),
                (1          ,"John"             ,1020, 52),
                (2          ,"Nikhil"           ,1021, 75),
                (3          ,"Stefen"           ,1025, 85),
                (4          ,"Thomas"           ,1026, 54),
                (5          ,"Ali"              ,1027, 87),
                (6          ,"Sanjay"           ,1028, 0)
            )
    (bottomRow, rightCol) = xlFile.setrange("Sheet1", 5,2,data)
    xlFile.addcomment("sheet1", "C11", "Absent")
    chrt1 = xlFile.insertchart("sheet2", 100, 100, 400, 200)
    xlFile.plotdata(sheet="sheet1",dataRanges=((6,3,bottomRow,5),),chartObject=chrt1,
                    title="Annual Examination : History", plotBy=2,categoryLabels=1,
                    seriesLabels=0, chartType="Bar")
    #~ xlFile.clearrange("sheet1",(3,2,3,5),"y")
    #~ xlFile.addcomment("sheet1", "B4", "Test Comment")
    #~ chrt1 = xlFile.insertchart("sheet1", 100, 100, 400, 250)
    #~ xlFile.plotdata(sheet="sheet1",dataRange=(4,2,bottomRow,rightCol), chartObject=chrt1, title="Test Chart", chartType="Column")
    #~ xlFile.copyrange(("sheet1","C3:E3"), ("sheet2", "C3"))
    #~ chrt2 = xlFile.insertchart("sheet2", 100, 100, 400, 250)
    #~ xlFile.movechart(chrt1,chrt2)
    #~ xlFile.copychart(chrt1,("sheet3","D22"), "y")
    #~ xlFile.hiderow("sheet1",7)
    #~ print xlFile.excelfunction("sheet1", (3,2,3,5), "Min")
    #~ print xlFile.getrange("sheet1","A2","C3")
    #~ xlFile.setcellfont("sheet1","Regular", "A1")
    #~ cellVal1 = xlFile.getcell("sheet1","A1")
    #~ xlFile.setcell("sheet1", cellVal1,1,3)
    #~ xlFile.setcellfont("sheet1","bold","C1")
    #~ xlFile.setcellfont("sheet1","italic",1,3)
    #~ xlFile.setcellfont("sheet1","underline",1,3)
    #~ xlFile.setcellalign("sheet1","left",1,3)
    #~ print xlFile.getrange("sheet1", "D5", "F6")
    #~ xlFile.setrange("sheet1", 10,10,((45,67),(67,"342"),(88,66.8),(32,77),(3,3)))
    #~ xlFile.addnewworksheetafter("sheet1", "Srijit1")
#-------------------------------------------------------------------------------------------------------------


if (__name__ == "__main__"):
    excelapp()
