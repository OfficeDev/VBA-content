---
title: Create Address Labels from a Spreadsheet
ms.prod: excel
ms.assetid: 6c08634c-8137-9c27-f4de-390a2b8ffb4c
ms.date: 06/08/2017
---


# Create Address Labels from a Spreadsheet




 **Sample code provided by:** Bill Jelen, [MrExcel.com](http://www.mrexcel.com/)



The following code example takes a list of addresses arranged as one address per row, and copies them onto another sheet, arranging them to fit on printable address labels. The workbook must contain two sheets, one named "Addresses" and one named "Labels". The addresses on the Addresses sheet must be arranged as one address per row, with the Name in Column A, Address Line 1 in Column B, Address Line 2 in Column C, and the City, State, Country/Region and Postal code in Column D. The addresses are rearranged and copied onto the Labels sheet.



```vb
Sub CreateLabels() 
    ' Clear out all records on Labels 
    Dim LabelSheet As Worksheet 
    Set LabelSheet = Worksheets("Labels") 
    LabelSheet.Cells.ClearContents 
 
    ' Set column width for labels 
    LabelSheet.Cells(1, 1).ColumnWidth = 35 
    LabelSheet.Cells(1, 2).ColumnWidth = 36 
    LabelSheet.Cells(1, 3).ColumnWidth = 30 
     
    ' Loop through all records 
    Dim AddressSheet As Worksheet 
    Set AddressSheet = Worksheets("Addresses") 
    FinalRow = AddressSheet.Cells(65536, 1).End(xlUp).Row 
     
    If FinalRow > 1 Then 
        NextRow = 1 
        NextCol = 1 
        For i = 2 To FinalRow 
            ' Set up row heights 
            If NextCol = 1 Then 
                LabelSheet.Cells(NextRow, 1).Resize(4, 1).RowHeight = 15.25 
                LabelSheet.Cells(NextRow + 4, 1).RowHeight = 13.25 
            End If 
         
            ' Put the Name in row 1 
            ThisRow = NextRow 
            LabelSheet.Cells(ThisRow, NextCol).Value = AddressSheet.Cells(i, 1) &; "   " &; AddressSheet.Cells(i, 7) 
             
            ' Put the Address Line 1 in row 2 
            If AddressSheet.Cells(i, 2).Value > "" Then 
                ThisRow = ThisRow + 1 
                LabelSheet.Cells(ThisRow, NextCol).Value = AddressSheet.Cells(i, 2) 
            End If 
             
            ' Put the Address Line 2 in row 3 
            If AddressSheet.Cells(i, 3).Value > "" Then 
                ThisRow = ThisRow + 1 
                LabelSheet.Cells(ThisRow, NextCol).Value = AddressSheet.Cells(i, 3) 
            End If 
             
            ' Put the City, State, Country/Region and Postal code in row 4 
            If AddressSheet.Cells(i, 4).Value > "" Then 
                CitySt = AddressSheet.Cells(i, 4) 
            End If 
            ThisRow = ThisRow + 1 
            LabelSheet.Cells(ThisRow, NextCol).Value = CitySt 
             
            ' Update the row and column for the next label 
            If NextCol = 1 Then 
                NextCol = 2 
            ElseIf NextCol = 2 Then 
                NextCol = 3 
            Else 
                NextCol = 1 
                NextRow = NextRow + 5 
            End If 
         
        Next i 
         
        LabelSheet.Activate 
    Else 
        MsgBox "No records match the criteria" 
    End If 
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 


