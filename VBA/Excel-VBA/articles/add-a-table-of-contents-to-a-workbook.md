---
title: Add a Table of Contents to a Workbook
ms.prod: excel
ms.assetid: fc61a9c1-d651-502a-c8d4-d6a570898191
ms.date: 06/08/2017
---


# Add a Table of Contents to a Workbook

The following examples show different approaches for adding a table of contents to an Excel workbook.


 **Sample code provided by:** Dennis Wallentin, [VSTO &; .NET &; Excel](http://xldennis.wordpress.com/)

This example uses the  [Pages.Count Property (Excel)](pages-count-property-excel.md) property to calculate the number of pages on each sheet. In addition, the entries in the TOC link to their respective sheets to improve on-screen workbook navigation.




```vb
Option Explicit 
Sub Create_TOC() 
Dim wbBook As Workbook 
Dim wsActive As Worksheet 
Dim wsSheet As Worksheet 
Dim lnRow As Long 
Dim lnPages As Long 
Dim lnCount As Long 
Set wbBook = ActiveWorkbook 
With Application 
    .DisplayAlerts = False 
    .ScreenUpdating = False 
End With 
'If the TOC sheet already exist delete it and add a new 
'worksheet. 
On Error Resume Next 
With wbBook 
    .Worksheets("TOC").Delete 
    .Worksheets.Add Before:=.Worksheets(1) 
End With 
On Error GoTo 0 
Set wsActive = wbBook.ActiveSheet 
With wsActive 
    .Name = "TOC" 
    With .Range("A1:B1") 
        .Value = VBA.Array("Table of Contents", "Sheet # - # of Pages") 
        .Font.Bold = True 
    End With 
End With 
lnRow = 2 
lnCount = 1 
'Iterate through the worksheets in the workbook and create 
'sheetnames, add hyperlink and count &; write the running number 
'of pages to be printed for each sheet on the TOC sheet. 
For Each wsSheet In wbBook.Worksheets 
    If wsSheet.Name <> wsActive.Name Then 
        wsSheet.Activate 
        With wsActive 
            .Hyperlinks.Add .Cells(lnRow, 1), "", _ 
            SubAddress:="'" &; wsSheet.Name &; "'!A1", _ 
            TextToDisplay:=wsSheet.Name 
            lnPages = wsSheet.PageSetup.Pages().Count 
            .Cells(lnRow, 2).Value = "'" &; lnCount &; "-" &; lnPages 
        End With 
        lnRow = lnRow + 1 
        lnCount = lnCount + 1 
    End If 
Next wsSheet 
wsActive.Activate 
wsActive.Columns("A:B").EntireColumn.AutoFit 
With Application 
    .DisplayAlerts = True 
    .ScreenUpdating = True 
End With 
End Sub
```

 **Sample code provided by:** Bill Jelen, [MrExcel.com](http://www.mrexcel.com/)
This example verifies that a sheet named "TOC" already exists. If it exists, the example updates the table of contents. Otherwise, the example creates a new TOC sheet at the beginning of the workbook. The name of each worksheet, along with the corresponding printed page numbers, is listed in the table of contents. To retrieve the page numbers the example opens the Print Preview dialog box. You must close the dialog box and then the table of contents is created.



```vb
Sub CreateTableOfContents() 
    ' Determine if there is already a Table of Contents 
    ' Assume it is there, and if it is not, it will raise an error 
    ' if the Err system variable is > 0, you know the sheet is not there 
    Dim WST As Worksheet 
    On Error Resume Next 
    Set WST = Worksheets("TOC") 
    If Not Err = 0 Then 
        ' The Table of contents doesn't exist. Add it 
        Set WST = Worksheets.Add(Before:=Worksheets(1)) 
        WST.Name = "TOC" 
    End If 
    On Error GoTo 0 
     
    ' Set up the table of contents page 
    WST.[A2] = "Table of Contents" 
    With WST.[A6] 
        .CurrentRegion.Clear 
        .Value = "Subject" 
    End With 
    WST.[B6] = "Page(s)" 
    WST.Range("A1:B1").ColumnWidth = Array(36, 12) 
    TOCRow = 7 
    PageCount = 0 
 
    ' Do a print preview on all sheets so Excel calcs page breaks 
    ' The user must manually close the PrintPreview window 
    Msg = "Excel needs to do a print preview to calculate the number of pages. " 
    Msg = Msg &; "Please dismiss the print preview by clicking close." 
    MsgBox Msg 
    ActiveWindow.SelectedSheets.PrintPreview 
 
    ' Loop through each sheet, collecting TOC information 
    For Each S In Worksheets 
        If S.Visible = -1 Then 
            S.Select 
            ThisName = ActiveSheet.Name 
            HPages = ActiveSheet.HPageBreaks.Count + 1 
            VPages = ActiveSheet.VPageBreaks.Count + 1 
            ThisPages = HPages * VPages 
 
            ' Enter info about this sheet on TOC 
            Sheets("TOC").Select 
            Range("A" &; TOCRow).Value = ThisName 
            Range("B" &; TOCRow).NumberFormat = "@" 
            If ThisPages = 1 Then 
                Range("B" &; TOCRow).Value = PageCount + 1 &; " " 
            Else 
                Range("B" &; TOCRow).Value = PageCount + 1 &; " - " &; PageCount + ThisPages 
            End If 
        PageCount = PageCount + ThisPages 
        TOCRow = TOCRow + 1 
        End If 
    Next S 
End Sub
```


## About the Contributors
<a name="AboutContributor"> </a>

Dennis Wallentin is the author of VSTO &; .NET &; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 


