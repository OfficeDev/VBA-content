---
title: Sort Worksheets Alphanumerically by Name
ms.prod: excel
ms.assetid: 20ec8072-4886-40bc-8784-ab3d100d613a
ms.date: 06/08/2017
---


# Sort Worksheets Alphanumerically by Name

The following example shows how to sort the worksheets in a workbook alphanumerically based on the name of the sheet by using the  **[Name](worksheet-name-property-excel.md)** property of the **[Worksheet](worksheet-object-excel.md)** object.

 **Sample code provided by:** Tom Urtis, [Atlas Programming Management](http://www.atlaspm.com/)



```vb
Sub SortSheetsTabName()
    Application.ScreenUpdating = False
    Dim iSheets%, i%, j%
    iSheets = Sheets.Count
    For i = 1 To iSheets - 1
        For j = i + 1 To iSheets
            If Sheets(j).Name < Sheets(i).Name Then
                Sheets(j).Move before:=Sheets(i)
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the coauthor of "Holy Macro! It's 2,500 Excel VBA Examples." 


