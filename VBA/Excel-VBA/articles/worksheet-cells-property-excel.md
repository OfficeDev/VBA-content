---
title: Worksheet.Cells Property (Excel)
keywords: vbaxl10.chm175080
f1_keywords:
- vbaxl10.chm175080
ms.prod: excel
api_name:
- Excel.Worksheet.Cells
ms.assetid: 19c14e41-7d8e-b56f-fd60-717df64edee8
ms.date: 06/08/2017
---


# Worksheet.Cells Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents all the cells on the worksheet (not just the cells that are currently in use).


## Syntax

 _expression_ . **Cells**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

Because the  **[Item](range-item-property-excel.md)** property is the default property for the **Range** object, you can specify the row and column index immediately after the **Cells** keyword. For more information, see the **Item** property and the examples for this topic.

Using this property without an object qualifier returns a  **Range** object that represents all the cells on the active worksheet.


## Example

This example sets the font size for cell C5 on Sheet1 to 14 points.


```vb
Worksheets("Sheet1").Cells(5, 3).Font.Size = 14
```

This example clears the formula in cell one on Sheet1.




```vb
Worksheets("Sheet1").Cells(1).ClearContents
```

This example sets the font and font size for every cell on Sheet1 to 8-point Arial




```vb
With Worksheets("Sheet1").Cells.Font 
    .Name = "Arial" 
    .Size = 8 
End With
```

 **Sample code provided by:** Tom Urtis,[Atlas Programming Management](http://www.atlaspm.com/)

This example toggles a sort between ascending and descending order when you double-click any cell in the data range. The data is sorted based on the column of the cell that is double-clicked.




```vb
Option Explicit
Public blnToggle As Boolean

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim LastColumn As Long, keyColumn As Long, LastRow As Long
    Dim SortRange As Range
    LastColumn = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    keyColumn = Target.Column
    
    If keyColumn <= LastColumn Then
    
        Application.ScreenUpdating = False
        Cancel = True
        LastRow = Cells(Rows.Count, keyColumn).End(xlUp).Row
        Set SortRange = Target.CurrentRegion
        
        blnToggle = Not blnToggle
        If blnToggle = True Then
            SortRange.Sort Key1:=Cells(2, keyColumn), Order1:=xlAscending, Header:=xlYes
        Else
            SortRange.Sort Key1:=Cells(2, keyColumn), Order1:=xlDescending, Header:=xlYes
        End If
    
        Set SortRange = Nothing
        Application.ScreenUpdating = True
        
    End If
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the co author of "Holy Macro! It's 2,500 Excel VBA Examples." 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

