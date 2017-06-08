---
title: Worksheet.HPageBreaks Property (Excel)
keywords: vbaxl10.chm175135
f1_keywords:
- vbaxl10.chm175135
ms.prod: excel
api_name:
- Excel.Worksheet.HPageBreaks
ms.assetid: 0d26aa71-714f-a6a0-8a10-4ea6bd7d852d
ms.date: 06/08/2017
---


# Worksheet.HPageBreaks Property (Excel)

Returns an  **[HPageBreaks](hpagebreaks-object-excel.md)** collection that represents the horizontal page breaks on the sheet. Read-only.


## Syntax

 _expression_ . **HPageBreaks**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

There is a limit of 1,026 horizontal page breaks per sheet.


## Example

The following code example displays the number of full-screen and print-area horizontal page breaks.


```vb
For Each pb in Worksheets(1).HPageBreaks 
    If pb.Extent = xlPageBreakFull Then 
        cFull = cFull + 1 
    Else 
        cPartial = cPartial + 1 
    End If 
Next 
MsgBox cFull &; " full-screen page breaks, " &; cPartial &; _ 
    " print-area page breaks"
```



 **Sample code provided by:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)

The following code example adds a page break when the value of a cell in column A changes.




```vb
Sub AddPageBreaks() 
    StartRow = 2 
    FinalRow = Range("A65536").End(xlUp).Row 
    LastVal = Cells(StartRow, 1).Value 
    For i = StartRow To FinalRow 
    ThisVal = Cells(i, 1).Value 
    If Not ThisVal = LastVal Then 
    ActiveSheet.HPageBreaks.Add before:=Cells(i, 1) 
    End If 
    LastVal = ThisVal 
    Next i 
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

