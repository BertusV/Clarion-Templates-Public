#Template(XLSXWriter,'XLSX Writer Class Templates'),Family('ABC','CW20')
#!------------------------------------------------------------------------------------!#
#! XLSX Writer Version 1.0
#! (c) 2018 by Bertus Viljoen
#! All rights reserved
#! This Template is written to implement the XLSX Writter class written by 
#! RaFael from https://libxlsxwriter.github.io 
#!
#!------------------------------------------------------------------------------------!#
#Extension(Activate_XLSXWriter,'Activate XLSX Writer'),Application
  #Display('XLSX Writer')
  #Display('(c) 2018 All Rights Reserved')
  #Display('Version 1.0')
  #Display()
  #Display('See "Adding XLSX Writer to your application."')
  #Display('This template implements RaFael''s XLSX Writer Class.')
  #Display()
  #Prompt('Disable XLSX Writer in this App',Check),%DisableXW,AT(10)
#!
#!------------------------------------------------------------------------------------!#
#AT (%AfterGlobalIncludes),where(%DisableXW=0)
  #Call(%XLSXWriterInclude)
#EndAt
#!
#AT(%CustomGlobalDeclarations),Where(%DisableXW=0)
    #PROJECT('None(libxlsxw.dll), CopyToOutputDirectory=Always')
    #PROJECT('libxlsxw.lib')
#EndAt    
#!------------------------------------------------------------------------------------!#
#!
#Extension(IncludeXLSXWriter,'XLSX Writer'),Description('XLSX Writer Class implementation'),Procedure,Req(Activate_XLSXWriter)
#Prepare
#Declare(%XWCellTypes_,'String|Number|Formula|Boolean|DateTime')
#EndPrepare
#Sheet
#Tab('General')
    #Boxed('XLSX Writer')
        #Prompt('XLSX Writer object name',@S25),%XLSX,DEFAULT('xlsx'),REQ
        #!Prompt('Tile Buttons',Control),%TileButtons,Multi('Tile Buttons'),REQ
        #Prompt('Disable XLSX Writer in this Procedure',Check),%DisableXWP,AT(10)
        #Prompt('Workbook Name:',@S100),%XLSXWorkbookName,REQ
        #Button('Sheets'),MULTI(%XLSXSheets,%XLSXSheetName),INLINE
            #Prompt('Sheet Names',@S25),%XLSXSheetName,REQ
            #Prompt('Create Sheet Automatically',CHECK),%XWCreateSheet,DEFAULT(%TRUE)
            #INSERT(%XWSheetDetails)
        #EndButton
        #Boxed('Workbook Properties')
            #Prompt('Title:',@S50),%XWTitle
            #Prompt('Author:',@S25),%XWAuthor
            #Prompt('Company:',@s50),%XWCompany
            #Prompt('Comments:',@s100),%XWComments
            #Display('Put literal strings in invereted commas')
        #EndBoxed
    #EndBoxed
#EndTab
#EndSheet
#!
#! Declare Local data
#AT (%DataSection,''),where(%DisableXW=0 AND %DisableXWP=0),Priority(8500)
#Declare(%XWCellColumn,LONG)
#Declare(%XWGraphDataStartColumnNumber,LONG)
#Declare(%XWGraphDataEndColumnNumber,LONG)
#Declare(%XWGraphCatStartColumnNumber,LONG)
#Declare(%XWGraphCatEndColumnNumber,LONG)
     #Call(%XLSXWriterLocalData)
#EndAt
#!
#Code(CreateXLSXWorkbook,'XLSX Writer create workbook'),Description('XLSX Writer create workbook'),REQ(IncludeXLSXWriter)
#Sheet
#Tab('General')
    #Display('No prompts needed')
#EndTab
#EndSheet
#!
#! Take the steps needed to create workbook and sheets
    #Call(%XLSXWriterCreateWorkbook)
#!
#Code(CreateXLSXWorkSheet,'XLSX Writer create worksheet'),Description('XLSX Writer create worksheet'),REQ(IncludeXLSXWriter)
#Sheet
#Tab('General')
    #Prompt('Sheet Name:',FROM(%XLSXSheets,,%XLSXSheetName)),%ThisWXSheet,REQ
#EndTab
#EndSheet
#!
#! Take the steps needed to a create sheet
    #FIX(%XLSXSheets,%ThisWXSheet)
    #FOR(%XLSXSheets)
        #If(%XLSXSheetName<>%ThisWXSheet)
            #Cycle
        #EndIf
        worksheet#=XW:%XLSX.AddWorksheet(%XLSXSheetName)
    #EndFor
#!
#Code(CloseXLSXWorkbook,'XLSX Writer close workbook'),Description('XLSX Writer close workbook'),REQ(IncludeXLSXWriter)
#Sheet
#Tab('General')
    #Display('No prompts needed')
#EndTab    
#EndSheet
#!
#! Take the steps needed to close off the workbook
    #Call(%XLSXWriterCloseWorkbook)
#!
#!
#Code(WriteXLSXWorkbookSheet,'XLSX Writer Write Sheet data workbook'),Description('XLSX Writer Write Sheet data workbook'),REQ(IncludeXLSXWriter)
#Sheet
#Tab('General')
    #Prompt('Sheet Name:',FROM(%XLSXSheets,,%XLSXSheetName)),%ThisWXSheet,REQ
    #Prompt('Write Details',DROP('Data[DATA]|Header[HEADER]|Totals[TOTALS]|Graph[GRAPH]')),%XWSheetWriteType,DEFAULT('DATA'),REQ
    #Boxed('Create Sheet'),WHERE(%XWSheetWriteType = 'HEADER')
    #Prompt('Create Sheet First',CHECK),%XWCreateSheet,DEFAULT(%False)
    #EndBoxed
#EndTab    
#EndSheet
#!
#! Write Either the data or the header
#!    #Call(%XWWriteData)

    #FIX(%XLSXSheets,%ThisWXSheet)
    #FOR(%XLSXSheets)
        #If(%XLSXSheetName<>%ThisWXSheet)
            #Cycle
        #EndIf
        #If(%XWCreateSheet=%True)
            worksheet#=XW:%XLSX.AddWorksheet(%XLSXSheetName)
        #Else
            err#=XW:%XLSX.GetWorksheetByName(%ThisWXSheet)
        #EndIf
        #SET(%XWCellColumn,%XWCellStartColumn)
            !%XWSheetWriteType
        #IF(%XWSheetWriteType='HEADER')
            #IF(%XWCellStartRow<>0)
            %XWCellRow = %XWCellStartRow
            #EndIf        
            #For(%XWCellValues)
                #Case(%XWCellHType)
                #Of('String')
                    #Call(%XLSXWriteCellHString)
                #Of('Number')
                    #Call(%XLSXWriteCellHNumber)
                #Of('Formula')
                    #Call(%XLSXWriteCellHFormula)
                #Of('Boolean')
                    #Call(%XLSXWriteCellHBoolean)
                #Of('DateTime')
                    #Call(%XLSXWriteCellHDateTime)
                #EndCase
                #Set(%XWCellColumn,%XWCellColumn + 1)
            #EndFor
        #ElsIf(%XWSheetWriteType='DATA')
            #FOR(%XWCellValues)
                #Case(%XWCellType)
                #Of('String')
                    #Call(%XLSXWriteCellString)
                #Of('Number')
                    #Call(%XLSXWriteCellNumber)
                #Of('Formula')
                    #Call(%XLSXWriteCellFormula)
                #Of('Boolean')
                    #Call(%XLSXWriteCellBoolean)
                #Of('DateTime')
                    #Call(%XLSXWriteCellDateTime)
                #EndCase
                #Set(%XWCellColumn,%XWCellColumn + 1)
            #EndFor
        #ElsIf(%XWSheetWriteType='TOTALS')    
                If(%XWCellRow-1 > %XWCellStartRow+1)
            #Set(%XWCellColumn,1)
            #FOR(%XWCellValues)
                #If (%DisableXW=0 AND %XWSumCell=%True)
                    #If(%XWCellBeginEnd=%True)
                    Begin
                    #EndIf
                    #Call(%XLSXWriteCellFormat)
                    #Call(%XLSXWriteCellFormatExcelMask)
                    err#=XW:%XLSX.WriteFormula(%XWCellRow,%XWCellColumn,'SUM('&XW:ColumnString[%XWCellColumn]&%XWCellStartRow+1&':'&XW:ColumnString[%XWCellColumn]&%XWCellRow-1&')')
                    #If(%XWCellBeginEnd=%True)
                    End
                    #EndIf
                #EndIf
                #Set(%XWCellColumn,%XWCellColumn + 1)
            #EndFor
                End
        #ElsIf(%XWSheetWriteType='GRAPH')  
                #Set(%XWCellColumn,1)
                #FOR(%XWCellValues)
                    #If(%XWGraphDataStartColumn = %XWCellValue)
                        #Set(%XWGraphDataStartColumnNumber,%XWCellColumn)
                    #EndIf
                    #If(%XWGraphDataEndColumn = %XWCellValue)
                        #Set(%XWGraphDataEndColumnNumber,%XWCellColumn)
                    #EndIf
                    #If(%XWGraphCatStartColumn = %XWCellValue)
                        #Set(%XWGraphCatStartColumnNumber,%XWCellColumn)
                    #EndIf
                    #If(%XWGraphCatEndColumn = %XWCellValue)
                        #Set(%XWGraphCatEndColumnNumber,%XWCellColumn)
                    #EndIf
                    #Set(%XWCellColumn,%XWCellColumn + 1)
                #EndFor
                XW:%XLSX:Chart.AddChart(XW:%XLSX,%XWSheetGraphType)
                XW:Series = XW:%XLSX:Chart.AddSeries()
                XW:%XLSX:Chart.SetValues(XW:Series,XW:%XLSX.ActiveWorkSheetName,%XWCellStartRow+1,%XWGraphDataStartColumnNumber,%XWCellRow-1,%XWGraphDataEndColumnNumber)
                XW:%XLSX:Chart.SetCategories(XW:Series,XW:%XLSX.ActiveWorkSheetName,%XWCellStartRow+1,%XWGraphCatStartColumnNumber,%XWCellRow-1,%XWGraphCatEndColumnNumber)
                XW:%XLSX:Chart.SetTitleName(%XWGraphTitle)
                #EMBED(%XW_Chart_Format,'Format the graph before it''s added to the sheet')
                err# = XW:%XLSX:Chart.SetPoints(XW:Series,%XWCellRow-1) 
                err# = XW:%XLSX.InsertChart(%XWCellRow+1,1,XW:%XLSX:Chart.Chart,,,%XWGraphXScale,%XWGraphYScale)
        #EndIf    
        #IF(%XWCellsIncreaseRow = %True)
                %XWCellRow += 1
        #EndIf
    #EndFor
#!
#! Write a variable to the sheet
#Code(WriteXLSXCells,'XLSX Writer write Cells'),Description('XLSX Writer Write Cells'),REQ(IncludeXLSXWriter)
#Sheet
#Tab('General')
    #Prompt('Sheet Name:',FROM(%XLSXSheets,,%XLSXSheetName)),%ThisXLSXSheet,REQ
    #Prompt('Reset Row Counter',CHECK),%XWRestRowCounter
    #Button('Write Values'),MULTI(%XWCellValues,%XWCellValue),INLINE
#INSERT(%XWSheetValues)
    #EndButton
#EndTab    
#EndSheet
#!
    #FIX(%XLSXSheets,%ThisXLSXSheet)
            err#=XW:%XLSX.GetWorksheetByName(%ThisXLSXSheet)
    #FOR(%XLSXSheets)        
        #SET(%XWCellColumn,%XWCellStartColumn)
        #IF(%XWCellStartRow<>0 AND %XWRestRowCounter=%True)
                %XWCellRow = %XWCellStartRow
        #EndIf        
        #FOR(%XWCellValues)
            #Case(%XWCellType)
            #Of('String')
                #Call(%XLSXWriteCellString)
            #Of('Number')
                #Call(%XLSXWriteCellNumber)
            #Of('Formula')
                #Call(%XLSXWriteCellFormula)
            #Of('Boolean')
                #Call(%XLSXWriteCellBoolean)
            #Of('DateTime')
                #Call(%XLSXWriteCellDateTime)
            #EndCase
            #Set(%XWCellColumn,%XWCellColumn + 1)
        #EndFor
        #IF(%XWCellsIncreaseRow = %True)
                %XWCellRow += 1
        #EndIf
    #EndFor
#!    
#! Write a String to the sheet
#Code(WriteXLSXString,'XLSX Writer write String'),Description('XLSX Writer Write String'),REQ(IncludeXLSXWriter)
#Sheet
#Tab('General')
    #Prompt('String Name',@S50),%XWString
    #Prompt('Row',@S50),%XWStringRow
    #Prompt('Column',@S50),%XWStringColumn
    #Prompt('Reset Format',CHECK),%XWStringSetFormat
    #Prompt('Merge Cells',CHECK),%XWStringMergeCells
#EndTab    
#Tab('Format Cell'),WHERE(%XWStringSetFormat=%True)
    #Prompt('Font Name:',@s25),%XWStringFontName,PROP(PROP:FontName)
    #Prompt('Font Size:',@N2),%XWStringFontSize
    #Prompt('Font Colour:',DROP('Default[DEFAULT]|Black[COLOR:Black]|Maroon[COLOR:Maroon]|Green[COLOR:Green]|Olive[COLOR:Olive]|Orange[COLOR:Orange]|Navy[COLOR:Navy]|Purple[COLOR:Purple]|Teal[COLOR:Teal]|Gray[COLOR:Gray]|Silver[COLOR:Silver]|Red[COLOR:Red]|Lime[COLOR:Lime]|Yellow[COLOR:Yellow]|Blue[COLOR:Blue]|Fuchia[COLOR:Fuchsia]|Aqua[COLOR:Aqua]|White[COLOR:White]')),%XWStringFontColour
    #Prompt('Font Style:',DROP('Regular[FONT:regular]|Thin[FONT:thin]|Bold[FONT:bold]|Weight[FONT:weight]|Fixed[FONT:fixed]|Italic[FONT:italic]|Underline[FONT:underline]|Strikout[FONT:strikeout]')),%XWStringFontStyle
    #Prompt('Align Text:',DROP('None[XLSX:ALIGN_NONE]|Left[XLSX:ALIGN_LEFT]|Center[XLSX:ALIGN_CENTER]|Right[XLSX:ALIGN_RIGHT]|Fill[XLSX:ALIGN_FILL]|Justify[XLSX:ALIGN_JUSTIFY]|CenterAcross[XLSX:ALIGN_CENTER_ACROSS]|Distributed[XLSX:ALIGN_DISTRIBUTED]')),%XWStringAlignText,DEFAULT('')
    #Prompt('Align Vertical Text:',DROP('None[XLSX:ALIGN_NONE]|Top[XLSX:ALIGN_VERTICAL_TOP]Bottom[XLSX:ALIGN_VERTICAL_BOTTOM]|Center[XLSX:ALIGN_VERTICAL_CENTER]|Justify[XLSX:ALIGN_VERTICAL_JUSTIFY]|Distributed[XLSX:ALIGN_VERTICAL_DISTRIBUTED]')),%XWStringVAlignText,DEFAULT('')
    #Prompt('Cell Colour:',DROP('Default[DEFAULT]|Black[COLOR:Black]|Maroon[COLOR:Maroon]|Green[COLOR:Green]|Olive[COLOR:Olive]|Orange[COLOR:Orange]|Navy[COLOR:Navy]|Purple[COLOR:Purple]|Teal[COLOR:Teal]|Gray[COLOR:Gray]|Silver[COLOR:Silver]|Red[COLOR:Red]|Lime[COLOR:Lime]|Yellow[COLOR:Yellow]|Blue[COLOR:Blue]|Fuchia[COLOR:Fuchsia]|Aqua[COLOR:Aqua]|White[COLOR:White]')),%XWStringCellColour,DEFAULT('')
#EndTab    
#Tab('Merge Cells'),WHERE(%XWStringMergeCells=%True)
    #Prompt('Last Row to merge:',@S50),%XWStringRowMerge
    #Prompt('Last Column to merge:',@S50),%XWStringColumnMerge
#EndTab
#EndSheet
#!
    #Call(%XLSXWriteString)
#!
#!
#!
#!------------------------------------------------------------------------------------!#
#!                                                                                   		        !#
#! This Groups are part of the XLSX Writer template and Classes  	        !#
#!                                                                                    		        !#
#!------------------------------------------------------------------------------------!#
#!
#Group(%XWSheetDetails)
    #Boxed('Sheet information')
        #Prompt('In the same row:',Check),%XWCellsInRow
        #Prompt('Increase row number afterwards:',Check),%XWCellsIncreaseRow
        #Display('Only do when the row value is a variable.')
        #Prompt('Row:',@S50),%XWCellRow,REQ
        #Prompt('Starting Row:',@N2),%XWCellStartRow,DEFAULT(0)
        #Prompt('Stating Column:',@N2),%XWCellStartColumn,REQ,DEFAULT(1)
        #Prompt('Write Header after creation',Check),%XWWriteSheetHeaders
        #Prompt('Write Totals before close',Check),%XWWriteSheetTotals
        #Button('Write Values'),MULTI(%XWCellValues,%XWCellValue),INLINE
    #! Headers    
            #Insert(%XWSheetHeaders)
   #! Values
            #Insert(%XWSheetValues)
        #EndButton
        #Prompt('Generate a graph',Check),%XWSheetGraph
        #Boxed('Format Graph'),WHERE(%XWSheetGraph=%True)
            #Prompt('Graph Style',DROP('Area[XLSX:CHART_AREA]|Area Stacked[XLSX:CHART_AREA_STACKED]|Area Stacked Percent[XLSX:CHART_AREA_STACKED_PERCENT]|Bar[XLSX:CHART_BAR]|Bar Stacked[XLSX:CHART_BAR_STACKED]|Bar Stacked Percent[XLSX:CHART_BAR_STACKED_PERCENT]|Column[XLSX:CHART_COLUMN]|Column Stacked[XLSX:CHART_COLUMN_STACKED]|Column Stacked Percent[XLSX:CHART_COLUMN_STACKED_PERCENT]|Doughnut[XLSX:CHART_DOUGHNUT]|Line[XLSX:CHART_LINE]|Pie[XLSX:CHART_PIE]|Scatter[XLSX:CHART_SCATTER]|Scatter Straight[XLSX:CHART_SCATTER_STRAIGHT]|Scatter Straight with Markers[XLSX:CHART_SCATTER_STRAIGHT_WITH_MARKERS]|Scatter Smooth[XLSX:CHART_SCATTER_SMOOTH]|Scatter Smooth with Markers[XLSX:CHART_SCATTER_SMOOTH_WITH_MARKERS]|Radar[XLSX:CHART_RADAR]|Radar with Markers[XLSX:CHART_RADAR_WITH_MARKERS]|Radar Filled[XLSX:CHART_RADAR_FILLED]')),%XWSheetGraphType
            #Prompt('Graph Title',@S50),%XWGraphTitle,REQ
            #Prompt('Graph X Scale',@N4.1),%XWGraphXScale,REQ,DEFAULT(1)
            #Prompt('Graph Y Scale',@N4.1),%XWGraphYScale,REQ,DEFAULT(1)
            #Prompt('Categories Start Column',FROM(%XWCellValues,,%XWCellValue)),%XWGraphCatStartColumn,REQ #! FROM(%XWCellValues,,%XWCellValue)
            #Prompt('Categories End Column',FROM(%XWCellValues,,%XWCellValue)),%XWGraphCatEndColumn,REQ
            #Prompt('Data Start Column',FROM(%XWCellValues,,%XWCellValue)),%XWGraphDataStartColumn,REQ
            #Prompt('Data End Column',FROM(%XWCellValues,,%XWCellValue)),%XWGraphDataEndColumn,REQ
        #EndBoxed
     #EndBoxed
#!
#Group(%XWSheetHeaders)
            #Prompt('Heading:',@S50),%XWCellValueHeading,REQ
            #Prompt('Cell Type:',DROP('String|Number|Formula|Boolean|DateTime')),%XWCellHType
            #Prompt('Inclose in Begin/End:',CHECK),%XWCellHBeginEnd
            #Prompt('Reset Format:',CHECK),%XWCellHSetFormat
            #Boxed('Format Cell:'),WHERE(%XWCellHSetFormat=%True)
                #!Prompt('Font Name:',@s25),%XWCellHFontName
                #Prompt('Font Size:',@N2),%XWCellHFontSize
                #Prompt('Font Colour:',DROP('Default[DEFAULT]|Black[COLOR:Black]|Maroon[COLOR:Maroon]|Green[COLOR:Green]|Olive[COLOR:Olive]|Orange[COLOR:Orange]|Navy[COLOR:Navy]|Purple[COLOR:Purple]|Teal[COLOR:Teal]|Gray[COLOR:Gray]|Silver[COLOR:Silver]|Red[COLOR:Red]|Lime[COLOR:Lime]|Yellow[COLOR:Yellow]|Blue[COLOR:Blue]|Fuchia[COLOR:Fuchsia]|Aqua[COLOR:Aqua]|White[COLOR:White]')),%XWCellHFontColour
                #Prompt('Font Style:',DROP('Regular[FONT:regular]|Thin[FONT:thin]|Bold[FONT:bold]|Weight[FONT:weight]|Fixed[FONT:fixed]|Italic[FONT:italic]|Underline[FONT:underline]|Strikout[FONT:strikeout]')),%XWCellHFontStyle
                #Prompt('Align Text:',DROP('None[XLSX:ALIGN_NONE]|Left[XLSX:ALIGN_LEFT]|Center[XLSX:ALIGN_CENTER]|Right[XLSX:ALIGN_RIGHT]|Fill[XLSX:ALIGN_FILL]|Justify[XLSX:ALIGN_JUSTIFY]|CenterAcross[XLSX:ALIGN_CENTER_ACROSS]|Distributed[XLSX:ALIGN_DISTRIBUTED]')),%XWCellHAlignText,DEFAULT('')
                #Prompt('Align Vertical Text:',DROP('None[XLSX:ALIGN_NONE]|Top[XLSX:ALIGN_VERTICAL_TOP]Bottom[XLSX:ALIGN_VERTICAL_BOTTOM]|Center[XLSX:ALIGN_VERTICAL_CENTER]|Justify[XLSX:ALIGN_VERTICAL_JUSTIFY]|Distributed[XLSX:ALIGN_VERTICAL_DISTRIBUTED]')),%XWCellHVAlignText,DEFAULT('')
                #Prompt('Cell Colour:',DROP('Default[DEFAULT]|Black[COLOR:Black]|Maroon[COLOR:Maroon]|Green[COLOR:Green]|Olive[COLOR:Olive]|Orange[COLOR:Orange]|Navy[COLOR:Navy]|Purple[COLOR:Purple]|Teal[COLOR:Teal]|Gray[COLOR:Gray]|Silver[COLOR:Silver]|Red[COLOR:Red]|Lime[COLOR:Lime]|Yellow[COLOR:Yellow]|Blue[COLOR:Blue]|Fuchia[COLOR:Fuchsia]|Aqua[COLOR:Aqua]|White[COLOR:White]')),%XWCellHCellColour,DEFAULT('DEFAULT')
            #EndBoxed
            #Boxed('Format Number:'),WHERE(%XWCellHType='Number' OR %XWCellHType='Formula')
                #Prompt('Excel Mask:',@S25),%XWCellHExcelMask,DEFAULT('#,##0.00;-#,##0.00')
            #EndBoxed
            #Boxed('Time Datail'),WHERE(%XWCellHType='DateTime')
                #Prompt('Time:',@S25),%XWCellHTime
            #EndBoxed
#!
#!
#Group(%XWSheetValues)
            #Prompt('Value Name:',@S50),%XWCellValue,REQ
            #Prompt('Cell Type:',DROP('String|Number|Formula|Boolean|DateTime')),%XWCellType
            #Prompt('Inclose in Begin/End:',CHECK),%XWCellBeginEnd
            #Prompt('Reset Format:',CHECK),%XWCellSetFormat
            #Boxed('Format Cell:'),WHERE(%XWCellSetFormat=%True)
                #!Prompt('Font Name:',@s25),%XWCellFontName
                #Prompt('Font Size:',@N2),%XWCellFontSize
                #Prompt('Font Colour:',DROP('Default[DEFAULT]|Black[COLOR:Black]|Maroon[COLOR:Maroon]|Green[COLOR:Green]|Olive[COLOR:Olive]|Orange[COLOR:Orange]|Navy[COLOR:Navy]|Purple[COLOR:Purple]|Teal[COLOR:Teal]|Gray[COLOR:Gray]|Silver[COLOR:Silver]|Red[COLOR:Red]|Lime[COLOR:Lime]|Yellow[COLOR:Yellow]|Blue[COLOR:Blue]|Fuchia[COLOR:Fuchsia]|Aqua[COLOR:Aqua]|White[COLOR:White]')),%XWCellFontColour
                #Prompt('Font Style:',DROP('Regular[FONT:regular]|Thin[FONT:thin]|Bold[FONT:bold]|Weight[FONT:weight]|Fixed[FONT:fixed]|Italic[FONT:italic]|Underline[FONT:underline]|Strikout[FONT:strikeout]')),%XWCellFontStyle
                #Prompt('Align Text:',DROP('None[XLSX:ALIGN_NONE]|Left[XLSX:ALIGN_LEFT]|Center[XLSX:ALIGN_CENTER]|Right[XLSX:ALIGN_RIGHT]|Fill[XLSX:ALIGN_FILL]|Justify[XLSX:ALIGN_JUSTIFY]|CenterAcross[XLSX:ALIGN_CENTER_ACROSS]|Distributed[XLSX:ALIGN_DISTRIBUTED]')),%XWCellAlignText,DEFAULT('')
                #Prompt('Align Vertical Text:',DROP('None[XLSX:ALIGN_NONE]|Top[XLSX:ALIGN_VERTICAL_TOP]Bottom[XLSX:ALIGN_VERTICAL_BOTTOM]|Center[XLSX:ALIGN_VERTICAL_CENTER]|Justify[XLSX:ALIGN_VERTICAL_JUSTIFY]|Distributed[XLSX:ALIGN_VERTICAL_DISTRIBUTED]')),%XWCellVAlignText,DEFAULT('')
                #Prompt('Cell Colour:',DROP('Default[DEFAULT]|Black[COLOR:Black]|Maroon[COLOR:Maroon]|Green[COLOR:Green]|Olive[COLOR:Olive]|Orange[COLOR:Orange]|Navy[COLOR:Navy]|Purple[COLOR:Purple]|Teal[COLOR:Teal]|Gray[COLOR:Gray]|Silver[COLOR:Silver]|Red[COLOR:Red]|Lime[COLOR:Lime]|Yellow[COLOR:Yellow]|Blue[COLOR:Blue]|Fuchia[COLOR:Fuchsia]|Aqua[COLOR:Aqua]|White[COLOR:White]')),%XWCellCellColour,DEFAULT('DEFAULT')
            #EndBoxed
            #Boxed('Format Number:'),WHERE(%XWCellType='Number' OR %XWCellType='Formula')
                #Prompt('Excel Mask:',@S25),%XWCellExcelMask,DEFAULT('#,##0.00;-#,##0.00')
                #Prompt('Sum Column',CHECK),%XWSumCell,DEFAULT(%False)
            #EndBoxed
            #Boxed('Time Datail'),WHERE(%XWCellType='DateTime')
                #Prompt('Time:',@S25),%XWCellTime
            #EndBoxed
#!
#Group(%XLSXWriterInclude),Preserve
#If (%DisableXW=0)
    INCLUDE('XLSXWriter.INC'),ONCE
#EndIf
#!
#Group(%XLSXWriterLocalData),Preserve
#DECLARE(%XWRowCounterNumber,LONG)
#If (%DisableXW=0)
XW:%XLSX                       &xlsxwriter
XW:ColumnString                STRING('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
XW:FileName                     CString(256)
    #FOR(%XLSXSheets)
    #SET(%XWRowCounterNumber,%XWRowCounterNumber+1)
XW:RowCounter_%XWRowCounterNumber    LONG
    #EndFor
XW:%XLSX:Chart                 &xlsxchart
XW:Series                      LONG
#EndIf
#!
#Group(%XLSXWriterCreateWorkbook)
#If (%DisableXW=0)
    XW:%XLSX &= NEW(xlsxwriter)
    XW:%XLSX:Chart &= NEW(xlsxchart)
    XW:Filename = %XLSXWorkbookName
    XW:%XLSX.NewWorkBook(XW:FileName)
    #IF(%XWTitle<>'')
    XW:%XLSX.Properties.Title    = %XWTitle
    #EndIf
    #IF(%XWAuthor<>'')
    XW:%XLSX.Properties.Author   = %XWAuthor
    #EndIf
    #IF(%XWCompany<>'')
    XW:%XLSX.Properties.Company  = %XWCompany
    #EndIf
    #IF(%XWComments<>'')
    XW:%XLSX.Properties.Comments = %XWComments
    #EndIf
    err#=XW:%XLSX.SetProperties()
    err#=XW:%XLSX.SetCustomProperty('Creation date',,,,today())
#!   
    #FOR(%XLSXSheets)
        #IF(%XWCreateSheet=%True)
            #Call(%XLSXWriterCreateSheet)
        #EndIf    
    #EndFor
#!
#EndIf    
#!
#Group(%XLSXWriterCreateSheet)
#If (%DisableXW=0)
        worksheet#=XW:%XLSX.AddWorksheet(%XLSXSheetName)
        #IF(%XWCellStartRow<>0)
            %XWCellRow = %XWCellStartRow
        #EndIf        
        #SET(%XWCellColumn,%XWCellStartColumn)
        #IF(%XWWriteSheetHeaders=%True)
            #For(%XWCellValues)
                #Case(%XWCellHType)
                #Of('String')
                    #Call(%XLSXWriteCellHString)
                #Of('Number')
                    #Call(%XLSXWriteCellHNumber)
                #Of('Formula')
                    #Call(%XLSXWriteCellHFormula)
                #Of('Boolean')
                    #Call(%XLSXWriteCellHBoolean)
                #Of('DateTime')
                    #Call(%XLSXWriteCellHDateTime)
                #EndCase
                #Set(%XWCellColumn,%XWCellColumn + 1)
            #EndFor
            #IF(%XWCellsIncreaseRow = %True)
                %XWCellRow += 1
            #EndIf
        #EndIf
#EndIf    
#!
#Group(%XLSXWriterCloseWorkbook)
#If (%DisableXW=0)
        #FOR(%XLSXSheets)        
            #IF(%XWWriteSheetTotals=%True)
                err#=XW:%XLSX.GetWorksheetByName(%XLSXSheetName)
                #SET(%XWCellColumn,%XWCellStartColumn)
                #FOR(%XWCellValues)
                    #If (%DisableXW=0 AND %XWSumCell=%True)
                        #If(%XWCellBeginEnd=%True)
                    Begin
                        #EndIf
                        #Call(%XLSXWriteCellFormat)
                        #Call(%XLSXWriteCellFormatExcelMask)
                    err#=XW:%XLSX.WriteFormula(%XWCellRow,%XWCellColumn,'SUM('&XW:ColumnString[%XWCellColumn]&%XWCellStartRow+1&':'&XW:ColumnString[%XWCellColumn]&%XWCellRow-1&')')
                        #If(%XWCellBeginEnd=%True)
                    End
                        #EndIf
                    #EndIf
                    #Set(%XWCellColumn,%XWCellColumn + 1)
                #EndFor
                #IF(%XWCellsIncreaseRow = %True)
                    %XWCellRow += 1
                #EndIf
            #EndIf
        #EndFor
            err#=XW:%XLSX.CloseWorkbook()
            dispose(XW:%XLSX)
#EndIf
#!
#!
#GROUP(%XWWriteData)
    #FIX(%XLSXSheets,%ThisXLSXSheet)
            err#=XW:%XLSX.GetWorksheetByName(%ThisXLSXSheet)
    #IF(%XWCellStartRow<>0)
            %XWCellRow = %XWCellStartRow
    #EndIf        
    #SET(%XWCellColumn,%XWCellStartColumn)
    #FOR(%XWCellValues)
        #Case(%XWCellType)
        #Of('String')
            #Call(%XLSXWriteCellString)
        #Of('Number')
            #Call(%XLSXWriteCellNumber)
        #Of('Formula')
            #Call(%XLSXWriteCellFormula)
        #Of('Boolean')
            #Call(%XLSXWriteCellBoolean)
        #Of('DateTime')
            #Call(%XLSXWriteCellDateTime)
        #EndCase
        #Set(%XWCellColumn,%XWCellColumn + 1)
    #EndFor
    #IF(%XWCellsIncreaseRow = %True)
            %XWCellRow += 1
    #EndIf
#!
#!
#!
#Group(%XLSXWriteCellFormat)
#If (%DisableXW=0)
            #IF(%XWCellSetFormat=%True)
            XW:%XLSX.ClearFormat()
#!            #IF(%XWCellFontName<>'')
#!            XW:%XLSX.Format.Font=%XWCellFontName
#!            #EndIf
#!
            #IF(%XWCellCellColour<>'DEFAULT')
            XW:%XLSX.Format.Color=%XWCellCellColour
            #EndIf
#!
            #IF(%XWCellFontSize>0)
            XW:%XLSX.Format.FontSize=%XWCellFontSize
            #EndIf
#!
            #IF(%XWCellFontColour<>'DEFAULT')
            XW:%XLSX.Format.FontColor=%XWCellFontColour
            #EndIf
#!
            XW:%XLSX.Format.FontStyle=%XWCellFontStyle
#!
            #IF(%XWCellAlignText<>0)
            XW:%XLSX.Format.Align=%XWCellAlignText
            #EndIf
#!
            #IF(%XWCellVAlignText<>0)
            XW:%XLSX.Format.AlignV=%XWCellVAlignText
            #EndIf
#!
            #IF(%XWCellType='DateTime')
            XW:%XLSX.Format.Index=22
            XW:%XLSX.Format.Picture='@d6-'
            #EndIf
#!
            XW:%XLSX.SetFormat()
            #EndIf
#!
#EndIf
#!
#Group(%XLSXWriteCellFormatExcelMask)
#If (%DisableXW=0)
            #IF(%XWCellExcelMask<>'')
            XW:%XLSX.Format.ExcelMask='%XWCellExcelMask'
            XW:%XLSX.SetFormat()
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellString)
#If (%DisableXW=0)
            #If(%XWCellBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellFormat)
            err#=XW:%XLSX.WriteString(%XWCellRow,%XWCellColumn,%XWCellValue)
            #If(%XWCellBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellNumber)
#If (%DisableXW=0)
            #If(%XWCellBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellFormat)
            #Call(%XLSXWriteCellFormatExcelMask)
#!
            err#=XW:%XLSX.WriteNumber(%XWCellRow,%XWCellColumn,%XWCellValue)
            #If(%XWCellBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellFormula)
#If (%DisableXW=0)
            #If(%XWCellBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellFormat)
            #Call(%XLSXWriteCellFormatExcelMask)
#!
            err#=XW:%XLSX.WriteFormula(%XWCellRow,%XWCellColumn,%XWCellValue)
            #If(%XWCellBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellBoolean)
#If (%DisableXW=0)
            #If(%XWCellBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellFormat)
            err#=XW:%XLSX.WriteBoolean(%XWCellRow,%XWCellColumn,%XWCellValue)
            #If(%XWCellBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellDateTime)
#If (%DisableXW=0)
            #If(%XWCellBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellFormat)
            err#=XW:%XLSX.WriteDateTime(%XWCellRow,%XWCellColumn,%XWCellValue,%XWCellTime)
            #If(%XWCellBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellHFormat)
#If (%DisableXW=0)
            #IF(%XWCellHSetFormat=%True)
            XW:%XLSX.ClearFormat()
#!            #IF(%XWCellFontName<>'')
#!            XW:%XLSX.Format.Font=%XWCellFontName
#!            #EndIf
#!
            #IF(%XWCellHCellColour<>'DEFAULT')
            XW:%XLSX.Format.Color=%XWCellHCellColour
            #EndIf
#!
            #IF(%XWCellHFontSize>0)
            XW:%XLSX.Format.FontSize=%XWCellHFontSize
            #EndIf
#!
            #IF(%XWCellHFontColour<>'DEFAULT')
            XW:%XLSX.Format.FontColor=%XWCellHFontColour
            #EndIf
#!
            XW:%XLSX.Format.FontStyle=%XWCellHFontStyle
#!
            #IF(%XWCellHAlignText<>0)
            XW:%XLSX.Format.Align=%XWCellHAlignText
            #EndIf
#!
            #IF(%XWCellHVAlignText<>0)
            XW:%XLSX.Format.AlignV=%XWCellHVAlignText
            #EndIf
#!
            XW:%XLSX.SetFormat()
            #EndIf
#!
#EndIf
#!
#Group(%XLSXWriteCellHFormatExcelMask)
#If (%DisableXW=0)
            #IF(%XWCellHExcelMask<>'')
            XW:%XLSX.Format.ExcelMask='%XWCellHExcelMask'
            XW:%XLSX.SetFormat()
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellHString)
#If (%DisableXW=0)
            #If(%XWCellHBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellHFormat)
            err#=XW:%XLSX.WriteString(%XWCellRow,%XWCellColumn,%XWCellValueHeading)
            #If(%XWCellHBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellHNumber)
#If (%DisableXW=0)
            #If(%XWCellHBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellHFormat)
            #Call(%XLSXWriteCellHFormatExcelMask)
#!
            err#=XW:%XLSX.WriteNumber(%XWCellRow,%XWCellColumn,%XWCellValueHeading)
            #If(%XWCellHBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellHFormula)
#If (%DisableXW=0)
            #If(%XWCellHBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellHFormat)
            #Call(%XLSXWriteCellHFormatExcelMask)
#!
            err#=XW:%XLSX.WriteFormula(%XWCellRow,%XWCellColumn,%XWCellValueHeading)
            #If(%XWCellHBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellHBoolean)
#If (%DisableXW=0)
            #If(%XWCellHBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellHFormat)
            err#=XW:%XLSX.WriteBoolean(%XWCellRow,%XWCellColumn,%XWCellValueHeading)
            #If(%XWCellHBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteCellHDateTime)
#If (%DisableXW=0)
            #If(%XWCellHBeginEnd=%True)
            Begin
            #EndIf
            #Call(%XLSXWriteCellHFormat)
            XW:%XLSX.Format.Picture='@d6-'
            XW:%XLSX.SetFormat()
            err#=XW:%XLSX.WriteDateTime(%XWCellRow,%XWCellColumn,%XWCellValueHeading,%XWCellHTime)
            #If(%XWCellHBeginEnd=%True)
            End
            #EndIf
#EndIf
#!
#Group(%XLSXWriteString)
#If (%DisableXW=0)
            #IF(%XWStringSetFormat=%True)
            XW:%XLSX.ClearFormat()
            #IF(%XWStringFontName<>'')
            XW:%XLSX.Format.Font=%XWStringFontName
            #EndIf
#!
            #IF(%XWStringCellColour<>'DEFAULT')
            XW:%XLSX.Format.Color=%XWStringCellColour
            #EndIf
#!
            #IF(%XWStringFontSize>0)
            XW:%XLSX.Format.FontSize=%XWStringFontSize
            #EndIf
#!
            #IF(%XWStringFontColour<>'DEFAULT')
            XW:%XLSX.Format.FontColor=%XWStringFontColour
            #EndIf
#!
            XW:%XLSX.Format.FontStyle=%XWStringFontStyle
#!
            #IF(%XWStringAlignText<>0)
            XW:%XLSX.Format.Align=%XWStringAlignText
            #EndIf
#!
            #IF(%XWStringVAlignText<>0)
            XW:%XLSX.Format.AlignV=%XWStringVAlignText
            #EndIf
#!
            XW:%XLSX.SetFormat()
            #EndIf
#!
            #IF(%XWStringMergeCells=True)
            err#=XW:%XLSX.Merge(%XWStringRow,%XWStringColumn,%XWStringRowMerge,%XWStringColumnMerge)
            #EndIf
#!
            err#=XW:%XLSX.WriteString(%XWStringRow,%XWStringColumn,%XWString)
#EndIf
#!
