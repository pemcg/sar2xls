#!/usr/bin/env python
#
#--------------------------------------------------------------------------------------------------------------------------
# Author:	Peter McGowan
#               Copyright 2007 Peter McGowan 

#  This program is free software: you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation, either version 3 of the License, or
#  (at your option) any later version.

#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.

#  You should have received a copy of the GNU General Public License
#  along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
# Revision History
# Version:  0.1     06-Feb-07   PEMcG   Original version
#           0.2     27-Jun-09   PEMcG   Added CellDivisionFactor to ini file (eg convert Bytes -> KBytes etc)
#           0.3     28-Jun-09   PEMcG   Added the ability to add multiple data source columns together
#           0.4     29-Jun-09   PEMcG   Added the ability to plot multiple data source columns on one chart
#           0.5     02-Jul-09   PEMcG   Improved Error handling by liberal use of try/except
#           0.6     13-Jul-09   PEMcG   Removed "extraTitle" arg from NewxlFile.plotdata call for Excel 2007 compatibility
#                                       and added SaveFileName under new [General] section in the ini file
#           0.7     04-Aug-09   PEMcG   Added -v, -a and -D switches
#           0.8     12-May-11   PEMcG   Calculated lengths of columns rather than hard-coding (added MaxRow & MaxMaxRow)
#           0.9     01-Jun-11   PEMcG   Changed regex on "Performance Details for system:" search to allow '-' in hostname
#           0.10    24-Aug-11   PEMcG   Changed logic to handle the fact that not all books may contain all sheets
#           0.11    14-Sep-11   PEMcG   Added $SYSFQDN as a valid chart legent conversion ($SYSNAME translates to short name)
#           0.12    22-Feb-12   PEMcG   Edited to match sar2xls v4.0 format sheet titles
#           0.13    02-Mar-12   PEMcG   Renamed to GenGraphs
#           0.14    06-Mar-12   PEMcG   Use '::' rather than ":' as separators in Data section of the ini file
#
#
#--------------------------------------------------------------------------------------------------------------------------

Version = (0.14)

from UseExcel import UseExcel
from optparse import OptionParser
import ConfigParser, os, sys, re
import win32api
from win32com.client import constants
#
# The following list is used to convert a numeric column index into its character equivalent
#
LetterCorrespondingTo = ("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z")

usage = "usage: %prog [-f ini_file] [-a] [-V] [-D directory]"

parser = OptionParser(usage=usage)
parser.add_option("-f", dest="IniFilename",
                    help="specifies the ini file containing instructions")
parser.add_option("-V", "--version", action="store_true", dest="Version", default=False,
                    help="prints the version")
parser.add_option("-a", "--allfiles", action="store_true", dest="AllFiles", default=False,
                    help="process all ini files in the directory")
parser.add_option("-D", "--directory", dest="Directory", 
                    help="Specifies a directory to use for input and output")

(options, args) = parser.parse_args()
#
# Check the sanity of some of our arguments
#
# -v ?
#
if options.Version:
    print "GenGraphs.py version: " + str(Version)
    sys.exit()
    
if (options.IniFilename and options.AllFiles):
    parser.error("options -f and -a are mutually exclusive")
    parser.print_help()
    sys.exit()
  
if options.Directory:
    Directory = options.Directory
else:
    Directory = "."
#
# Now assemble the list of ini files to process (may be only one)
#
IniFileList = []
if options.IniFilename:
    #
    # We've been passed an ini file name, so just a single file for the list
    #
    IniFileList.append(Directory + "\\" + options.IniFilename)
elif options.AllFiles:
    #
    # Process all the ini files in the specified directory
    #
    for File in os.listdir(Directory):
        if os.path.splitext(File)[1] == ".ini":
            if not os.path.isdir(File):
                IniFileList.append(Directory + "\\" + File)
else:
    parser.error("must specify either -f or -a")
    parser.print_help()
    sys.exit()
#
# Check we can open the ini file
#
for IniFile in IniFileList:
    print "Processing " + IniFile + "..."
    try:
        File = open(IniFile, 'rb')
    except IOError:
        print "Can't open file " + IniFile + ", are you sure this file exists and is readable?"
        sys.exit()
    #
    # Seem to be ok, close it again
    #
    File.close()
    #
    # Read the [General] section from the inifile. If present, the SaveFileName parameter
    # tells us the Excel spreadsheet name to save the output file as
    #
    SaveFileName = False
    SaveFileNameList = win32api.GetProfileSection("General", IniFile)
    if len(SaveFileNameList) != 0:
        (SaveFileNameParameterName, SaveFileNameParameterValue) = SaveFileNameList[0].split("=")
        if re.match("SaveFileName", SaveFileNameParameterName):
            SaveFileName = SaveFileNameParameterValue
    #
    # Read the [Files] section from the inifile. This tells us the Excel spreadsheet
    # files that we are to read the source data from
    #
    FileList = win32api.GetProfileSection("Files", IniFile)
    if len(FileList) == 0:
        print "Could not read [Files] section from " + IniFile + ". Are you sure "\
        "you've specified an absolute path, not a relative path to the file?"
        sys.exit()
    #
    # Now split out the filename section (after the "=" character), and add all filenames to
    # the SourceFiles list
    #
    SourceFiles = []
    for File in FileList:
        SourceFiles.append(File.split("=")[1])
    #
    # At this point, SourceFiles is the list of input spreadsheet filenames
    #
    # -----
    #
    # Now read the [Charts] section from the inifile. This tells us how many charts
    # we will be plotting, and what the spreadsheet tab title will be for the chart data
    # and chart sheets in the new workbook
    #
    ChartListFromIniFile = win32api.GetProfileSection("Charts", IniFile)
    if len(ChartListFromIniFile) == 0:
        print "Could not read [Charts] section from " + IniFile
        sys.exit()
    #
    # Split out the chart number and sheettitle from the chart line (the bit after the "=" character)
    #
    Charts = {}
    for Chart in ChartListFromIniFile:
        (ChartNumber, SheetTitle) = Chart.split("=")
        Charts[ChartNumber] = SheetTitle
    #
    # At this point, Charts is a dictionary whose keys are the chart numbers, and whose values are
    # the sheet titles to be applied to the sheet tabs
    #
    # Now create a data structure called "ChartDetails"
    #
    # ChartDetails{Chartn}  {"SheetTitle" => String from [Charts]
    #                       {"GraphTitle" => String from [ChartnTitles]
    #                       {"YAxisTitle" => String from [ChartnTitles]
    #                       {"DataSourceSheets" => List[for each line under [ChartnData]]
    #                                                   {"NewColumnHeading" => String from [ChartnData]
    #                                                   {"CellDivisionFactor" => Integer from [ChartnData]
    #                                                   {"AddList"} => List[for each DataSource Sheet/Column to add together (could be only one)]
    #                                                           [{"SheetNameInSourceWorkbook"}
    #                                                           {"ColumnHeadingInSourceWorkbook"}]
    #
    # For example ChartDetails["Chart2"]["SheetTitle"] = "Sheet Title"
    #             ChartDetails["Chart2"]["GraphTitle"] = "Graph Title"
    #             ChartDetails["Chart2"]["YAxisTitle"] = "Y Axis title"
    #             ChartDetails["Chart2"]["DataSourceSheets"][0] = ["NewColumnHeading"] = "New Column heading"
    #                                                             ["CellDivisionFactor"] = 1024
    #                                                             ["AddList"][0] = ["SheetNameInSourceWorkbook"] = "Data Source Sheet Name"
    #                                                                              ["ColumnHeadingInSourceWorkbook"] = "Data Source Column Heading"
    #                                                                        [1] = ["SheetNameInSourceWorkbook"] = "Data Source Sheet Name"
    #                                                                              ["ColumnHeadingInSourceWorkbook"] = "Data Source Column Heading"
    #                                                       [1] = ["NewColumnHeading"] = "New Column heading"
    #                                                             ["CellDivisionFactor"] = 1
    #                                                             ["AddList"][0] = ["SheetNameInSourceWorkbook"] = "Data Source Sheet Name"
    #                                                                              ["ColumnHeadingInSourceWorkbook"] = "Data Source Column Heading"
    #                      
    ChartDetails = {}
    for Chart in Charts.keys():
        #
        # Read the [ChartnTitles] sections from the inifile. These tells us the main graph title
        # and Y Axis titles to be applied to each graph
        #
        ChartTitleListFromIniFile = win32api.GetProfileSection(Chart + "Titles", IniFile)
        if len(ChartTitleListFromIniFile) == 0:
            print "Could not read [ChartnTitles] section from " + IniFile
            sys.exit()
        ChartDetails[Chart] = {}
        ChartDetails[Chart]["SheetTitle"] = Charts[Chart]
        ChartDetails[Chart]["GraphTitle"] = ChartTitleListFromIniFile[0].split("=")[1]
        ChartDetails[Chart]["YAxisTitle"] = ChartTitleListFromIniFile[1].split("=")[1]
        
    for Chart in Charts.keys():
        #
        # Read the [ChartnData] sections from the inifile. These tell us the column heading to be applied
        # to each new column, the source sheet name to get the data from, and the source column heading on
        # that sheet
        #    
        ChartDataLinesFromIniFile = win32api.GetProfileSection(Chart + "Data", IniFile)
        ChartDetails[Chart]["DataSourceSheets"] = []
        for ChartDataLine in ChartDataLinesFromIniFile:
            (LeftHalfOfChartDataLine, RightHalfOfChartDataLine) = ChartDataLine.split("=")
            #
            # First deal with LeftHalfofChartDataLine
            # Set a default CellDivisionFactor in case there isn't one in the ini file
            #
            CellDivisionFactor = 1
            if re.search("::", LeftHalfOfChartDataLine):
                (NewColumnHeading, CellDivisionFactor) = re.split("::", LeftHalfOfChartDataLine)
            else:
                NewColumnHeading = LeftHalfOfChartDataLine
            Temp1 = {}
            Temp1["NewColumnHeading"] = NewColumnHeading
            Temp1["CellDivisionFactor"] = int(CellDivisionFactor)
            Temp1["AddList"] = []
            #
            # Deal with RightHalfofChartDataLine
            # See if we're adding values from multiple columns
            #
            match = re.search("Add\((.+)\)", RightHalfOfChartDataLine)
            if match:
                ChartDataLinesToAddTogether = match.group(1).split(",")
                for ChartDataLineToAdd in ChartDataLinesToAddTogether:
                    (SheetNameInSourceWorkbook, ColumnHeadingInSourceWorkbook) = ChartDataLineToAdd.split("::")
                    Temp2 = {}
                    Temp2["SheetNameInSourceWorkbook"] = SheetNameInSourceWorkbook
                    Temp2["ColumnHeadingInSourceWorkbook"] = ColumnHeadingInSourceWorkbook
                    Temp1["AddList"].append(Temp2)
            else:
                (SheetNameInSourceWorkbook, ColumnHeadingInSourceWorkbook) = RightHalfOfChartDataLine.split("::")
                Temp2 = {}
                Temp2["SheetNameInSourceWorkbook"] = SheetNameInSourceWorkbook
                Temp2["ColumnHeadingInSourceWorkbook"] = ColumnHeadingInSourceWorkbook
                Temp1["AddList"].append(Temp2)
            ChartDetails[Chart]["DataSourceSheets"].append(Temp1)
        # End for ChartDataLine in ChartDataLinesFromIniFile:
    
    #
    # Create a new spreadsheet
    #
    NewxlFile = UseExcel()
    NewxlFile.show()
    #
    # Open the source speadsheets
    #
    SourceBooks = []
    MaxMaxRow = 0
    for SourceFile in SourceFiles:
        try:
            xlBook = UseExcel(SourceFile)
        except:
            print "Can't open workbook " + SourceFile + ", are you sure this file exists?"
            continue
        Temp = {}
        Temp["FileName"] = SourceFile
        Temp["ExcelObject"] = xlBook
        #
        # Find out maximum row number - we're assuming that all sheets in this book have the same MaxRow as the first sheet
        #
        Chart1SheetNameInFirstSourceWorkbook = ChartDetails["Chart1"]["DataSourceSheets"][0]["AddList"][0]["SheetNameInSourceWorkbook"]
        Temp["MaxRow"] = xlBook.lastcellincolumn(Chart1SheetNameInFirstSourceWorkbook, "A")
        if Temp["MaxRow"] > MaxMaxRow:
            MaxMaxRow = Temp["MaxRow"]
        SourceBooks.append(Temp)
    #
    # Bail out to the next ini file if we have no spreadsheets to process
    #
    if len(SourceBooks) == 0:
        continue
    #
    # SourceBooks is now a list of Dictionary objects corresponding to the input spreadsheet files
    #
    #   i.e.    SourceBooks[0]["FileName"] = "C:\sar_files\DR\June\sysora1-2009-06-22.xls"
    #                         ["ExcelObject"] = UseExcel Object
    #           SourceBooks[1]["FileName"] = "C:\sar_files\DR\June\sysora2-2009-06-22.xls"
    #                         ["ExcelObject"] = UseExcel Object
    #
    ThisChart = {}
    HeaderRow = []
    #
    # Frig around to process the chartdetails keys by order of Chartn number
    #
    AllCharts = ChartDetails.keys()
    TempDictionary = {}
    for ThisChart in AllCharts:
        #
        # Split out the chart number and stick add it to the "TempDictionary" dictionary
        #
        TempDictionary[int(re.search("Chart(\d+)", ThisChart).group(1))] = ThisChart
    #
    # Sort the keys of this dictionary as a list called "TempKeys"
    #
    TempKeys = TempDictionary.keys()
    TempKeys.sort(lambda x, y: x-y)
    #
    # End of frig
    #
    # Now iterate through this (frigged) sorted list
    #
    for TempKey in TempKeys:
        Chart = TempDictionary[TempKey]
        ThisChart = ChartDetails[Chart]
    # for ThisChart in ChartDetails.keys()
        #
        # Create and Name the sheet
        #
        NewxlFile.addnewworksheetafter("Sheet1", ThisChart["SheetTitle"])
        #
        # Copy the time column from the first source book that has the required sheet
        #
        CellRange = "A1:A1"
        for ThisBook in range(len(SourceBooks)):
            try:
                SourceBooks[ThisBook]["ExcelObject"].copycolumntoclipboard((ThisChart["DataSourceSheets"][0]["AddList"][0]["SheetNameInSourceWorkbook"], CellRange))
            except:
                print "Could not find sheet name \"" + ThisChart["DataSourceSheets"][0]["AddList"][0]["SheetNameInSourceWorkbook"] \
                    + "\" in file \"" + SourceBooks[ThisBook]["FileName"] + "\". Check spelling."
                continue
            break
        NewxlFile.pasterangefromclipboard((ThisChart["SheetTitle"], "A1"))
        
        NewSheetColumn = 0
        #
        # Now iterate through SourceBooks (the list of open Excel data source workbooks)
        #
        for ThisBook in range(len(SourceBooks)):
            #
            # Pull out the system short name & FQDN from the overview page
            #
            TempValue = SourceBooks[ThisBook]["ExcelObject"].getcell("Overview", "A2")
            SystemName = re.search("Performance Details for system: ([\w-]+)", TempValue).group(1)
            SystemFQDN = re.search("Performance Details for system: ([\w-]+\.?.+)", TempValue).group(1)
            #
            # Pull out the sar date from the overview page
            #
            TempValue = SourceBooks[ThisBook]["ExcelObject"].getcell("Overview", "A6")
            SarDate = re.search("Statistics for (\d+-\d\d-\d\d)", TempValue).group(1)
            #
            # Iterate through our DataSourceSheets
            #
            for DataSourceSheet in ThisChart["DataSourceSheets"]:
                NewSheetColumn += 1
                #
                # Are we adding data from multiple source sheet columns?
                #
                if len(DataSourceSheet["AddList"]) > 1:
                    #
                    # Create/define a list of lists to hold our total values
                    #
                    TotalColumn = range(SourceBooks[ThisBook]["MaxRow"] - 2)
                    for i in range(len(TotalColumn)):
                        TotalColumn[i] = [0,]
                    #
                    # Now iterate through the source sheets and pull out the data
                    #
                    for SourceSheetNumber in range(len(DataSourceSheet["AddList"])):
                        #
                        # Read the header row from the sheet
                        #
                        CellRange = "A1:Z1"
                        try:
                            HeaderRow = SourceBooks[ThisBook]["ExcelObject"].getrange(DataSourceSheet["AddList"][SourceSheetNumber]["SheetNameInSourceWorkbook"], CellRange)
                        except:
                            print "Could not get range \"" + CellRange + "\" from sheet name \"" + DataSourceSheet["AddList"][SourceSheetNumber]["SheetNameInSourceWorkbook"] \
                                + "\" in file \"" + SourceBooks[ThisBook]["FileName"] + "\". Check spelling."
                            continue
                        #
                        # Get the column letter corresponding to the heading we're interested in
                        #
                        DataSourceColumn = LetterCorrespondingTo[list(HeaderRow[0]).index(DataSourceSheet["AddList"][SourceSheetNumber]["ColumnHeadingInSourceWorkbook"])]
                        #
                        # Now get the data column(s). 
                        #
                        CellRange = DataSourceColumn + "3:" + DataSourceColumn + str(SourceBooks[ThisBook]["MaxRow"])
                        try:
                            TempColumn = SourceBooks[ThisBook]["ExcelObject"].getrange(DataSourceSheet["AddList"][SourceSheetNumber]["SheetNameInSourceWorkbook"], CellRange)
                        except:
                            print "Could not get range \"" + CellRange + "\" from sheet name \"" + DataSourceSheet["AddList"][SourceSheetNumber]["SheetNameInSourceWorkbook"] \
                                + "\" in file \"" + SourceBooks[ThisBook]["FileName"] + "\". Check spelling."
                            sys.exit()
                        for Row in range(len(TempColumn)):
                            try:
                                TotalColumn[Row][0] += TempColumn[Row][0]
                            except TypeError:
                                #
                                # Column is shorter than usual, bail out of this loop
                                #
                                break
                    #
                    # Do we have a CellDivisionFactor to take into account (like convert Bytes - KBytes etc)?
                    #
                    if DataSourceSheet["CellDivisionFactor"] != 1:
                        #
                        # Apply the division factor to each cell
                        #
                        for Row in range(len(TotalColumn)):
                            TotalColumn[Row][0] = TotalColumn[Row][0] / DataSourceSheet["CellDivisionFactor"]
                    #
                    # Now write the total column back out the the new sheet
                    #
                    NewxlFile.setrange(ThisChart["SheetTitle"], 3, NewSheetColumn + 1, TotalColumn)
                else:
                    #
                    # Read the entire row of headers from the sheet
                    #
                    CellRange = "A1:Z1"
                    try:
                        HeaderRow = SourceBooks[ThisBook]["ExcelObject"].getrange(DataSourceSheet["AddList"][0]["SheetNameInSourceWorkbook"], CellRange)
                    except:
                        print "Could not get range \"" + CellRange + "\" from sheet name \"" + DataSourceSheet["AddList"][0]["SheetNameInSourceWorkbook"] \
                            + "\" in file \"" + SourceBooks[ThisBook]["FileName"] + "\". Check spelling."
                        continue
                    #
                    # Get the column letter corresponding to the heading we're interested in
                    #
                    try:
                        DataSourceColumn = LetterCorrespondingTo[list(HeaderRow[0]).index(DataSourceSheet["AddList"][0]["ColumnHeadingInSourceWorkbook"])]
                    except ValueError:
                        print "Could not find heading \"" + DataSourceSheet["AddList"][0]["ColumnHeadingInSourceWorkbook"] + "\" in sheet \"" \
                            + DataSourceSheet["AddList"][0]["SheetNameInSourceWorkbook"] + "\" in file \"" + SourceBooks[ThisBook]["FileName"] + "\". Check spelling."
                        sys.exit()
                    if DataSourceSheet["CellDivisionFactor"] != 1:
                        NewColumnData = []
                        #
                        # Read the column of data from the source sheet
                        #
                        CellRange = DataSourceColumn + "3:" + DataSourceColumn + str(SourceBooks[ThisBook]["MaxRow"])
                        try:
                            OldColumnData = SourceBooks[ThisBook]["ExcelObject"].getrange(DataSourceSheet["AddList"][0]["SheetNameInSourceWorkbook"], CellRange)
                        except:
                            print "Could not get range \"" + CellRange + "\" from sheet name \"" + DataSourceSheet["AddList"][0]["SheetNameInSourceWorkbook"] \
                                + "\" in file \"" + SourceBooks[ThisBook]["FileName"] + "\". Check spelling."
                            sys.exit()
                        #
                        # Apply the division factor to each cell
                        #
                        for Cell in OldColumnData:
                            try:
                                NewColumnData.append((Cell[0] / DataSourceSheet["CellDivisionFactor"],))
                            except TypeError:
                                #
                                # Column is shorter than usual, bail out of this loop
                                #
                                break
                        #
                        # Then write the amended column back out
                        #
                        NewxlFile.setrange(ThisChart["SheetTitle"], 3, NewSheetColumn + 1, NewColumnData)
                    else:
                        CellRange = DataSourceColumn + "1:" + DataSourceColumn + "1"
                        try:
                            SourceBooks[ThisBook]["ExcelObject"].copycolumntoclipboard((DataSourceSheet["AddList"][0]["SheetNameInSourceWorkbook"], CellRange))
                        except:
                            print "Could not copy column \"" + CellRange + "\" from sheet \"" + DataSourceSheet["AddList"][0]["SheetNameInSourceWorkbook"] \
                                + "\" in file \"" + SourceBooks[ThisBook]["FileName"] + "\". Check spelling."
                            sys.exit()
                        NewxlFile.pasterangefromclipboard((ThisChart["SheetTitle"], LetterCorrespondingTo[NewSheetColumn] + "1"))
                    # end if DataSourceSheet["CellDivisionFactor"] != 1:
                #
                # convert $SYSNAME, $SYSFQDN or $SARDATE in column heading to real system name or sar date
                #
                SysNameRegEx = re.compile("\$SYSNAME")
                SysFQDNRegEx = re.compile("\$SYSFQDN")
                SarDateRegEx = re.compile("\$SARDATE")
        
                NewColumnHeading = DataSourceSheet["NewColumnHeading"]
                if SysNameRegEx.search(NewColumnHeading):
                    NewColumnHeading = SysNameRegEx.sub(SystemName, NewColumnHeading)
                if SysFQDNRegEx.search(NewColumnHeading):
                    NewColumnHeading = SysFQDNRegEx.sub(SystemFQDN, NewColumnHeading)
                if SarDateRegEx.search(NewColumnHeading):
                    NewColumnHeading = SarDateRegEx.sub(SarDate, NewColumnHeading)
                #
                # Make the heading Bold, Arial, 10 pt
                #
                NewxlFile.setcellvalue(ThisChart["SheetTitle"], NewColumnHeading, LetterCorrespondingTo[NewSheetColumn] + "1", "Bold", "Arial", 10)
            # end for DataSourceSheet in ThisChart["DataSourceSheets"]:
        # end for SourceBook in SourceBooks:   
        
        #
        # Now add a chart
        #
        NewxlFile.addnewworksheetbefore(ThisChart["SheetTitle"], "Chart - " + ThisChart["SheetTitle"])
        NewChart = NewxlFile.insertchart("Chart - " + ThisChart["SheetTitle"], 1, 1, 900, 600)
        NewxlFile.plotdata(sheet = ThisChart["SheetTitle"],
                           dataRanges = ((1, 1, MaxMaxRow, NewSheetColumn + 1),),
                           chartObject = NewChart,
                           gallery = constants.xlLine,
                           format = 2,      # 2D line, unstacked with no markers
                           plotBy = constants.xlColumns,
                           categoryLabels = 1,
                           seriesLabels = 1,
                           hasLegend = True,
                           title = ThisChart["GraphTitle"],
                           categoryTitle = "Time",
                           valueTitle = ThisChart["YAxisTitle"],
                           extraTitle = ""
                           )
    #
    # Now close the source workbooks
    #
    for ThisBook in range(len(SourceBooks)):
        SourceBooks[ThisBook]["ExcelObject"].close()
    #
    # and tidy up by deleting the first sheet in the new workbook
    #
    NewxlFile.deleteworksheet("Sheet1")
    #
    # Save the file if we have a SaveFileName
    #
    if SaveFileName:
        NewxlFile.save(SaveFileName)
        NewxlFile.close()
        NewxlFile = None
# End for IniFile in IniFileList