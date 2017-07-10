---
title: Exporting a Table to a Word Document
ms.prod: excel
ms.assetid: 56ad67de-6f8b-4a55-a29e-4c2b5c88dfd5
ms.date: 06/08/2017
---

# Exporting a Table to a Word Document

This example takes the table named &;quot;Table1&;quot; on Sheet 1 and copies it into an existing Word document named &;quot;Quarter Report&;quot; at the bookmarked location named &;quot;Report&;quot;.

 **Sample code provided by:** Dennis Wallentin,[VSTO &; .NET &; Excel](http://xldennis.wordpress.com/)



```
Sub Export_Table_Word()

    'Name of the existing Word doc.
    Const stWordReport As String = "Quarter Report.docx"
    
    'Word objects.
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim wdbmRange As Word.Range
    
    'Excel objects.
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim rnReport As Range
    
    'Initialize the Excel objects.
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Sheet1")
    Set rnReport = wsSheet.Range("Table1")
    
    'Initialize the Word objets.
    Set wdApp = New Word.Application
    Set wdDoc = wdApp.Documents.Open(wbBook.Path &; "\" &; stWordReport)
    Set wdbmRange = wdDoc.Bookmarks("Report").Range
    
    'If the macro has been run before, clean up any artifacts before trying to paste the table in again.
    On Error Resume Next
    With wdDoc.InlineShapes(1)
        .Select
        .Delete
    End With
    On Error GoTo 0
    
    'Turn off screen updating.
    Application.ScreenUpdating = False
    
    'Copy the report to the clipboard.
    rnReport.Copy
    
    'Select the range defined by the "Report" bookmark and paste in the report from clipboard.
    With wdbmRange
        .Select
        .PasteSpecial Link:=False, _
                      DataType:=wdPasteMetafilePicture, _
                      Placement:=wdInLine, _
                      DisplayAsIcon:=False
    End With
    
    'Save and close the Word doc.
    With wdDoc
        .Save
        .Close
    End With
    
    'Quit Word.
    wdApp.Quit
    
    'Null out your variables.
    Set wdbmRange = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    'Clear out the clipboard, and turn screen updating back on.
    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
    End With
    
    MsgBox "The report has successfully been " &; vbNewLine &; _
           "transferred to " &; stWordReport, vbInformation

End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Dennis Wallentin is the author of VSTO &; .NET &; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 


