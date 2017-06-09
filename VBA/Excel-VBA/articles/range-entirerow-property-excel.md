---
title: Range.EntireRow Property (Excel)
keywords: vbaxl10.chm144123
f1_keywords:
- vbaxl10.chm144123
ms.prod: excel
api_name:
- Excel.Range.EntireRow
ms.assetid: 9e66da51-6cef-4109-ea4e-2acaad42aa1f
ms.date: 06/08/2017
---


# Range.EntireRow Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the entire row (or rows) that contains the specified range. Read-only.


## Syntax

 _expression_ . **EntireRow**

 _expression_ A variable that represents a **Range** object.


## Example

This example sets the value of the first cell in the row that contains the active cell. The example must be run from a worksheet.


```vb
ActiveCell.EntireRow.Cells(1, 1).Value = 5
```

 **Sample code provided by:** Holy Macro! Books,[Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&;p=1)

This example sorts all the rows on a worksheet, including hidden rows.




```vb
Sub SortAll()
    'Turn off screen updating, and define your variables.
    Application.ScreenUpdating = False
    Dim lngLastRow As Long, lngRow As Long
    Dim rngHidden As Range
    
    'Determine the number of rows in your sheet, and add the header row to the hidden range variable.
    lngLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set rngHidden = Rows(1)
    
    'For each row in the list, if the row is hidden add that that row to the hidden range variable.
    For lngRow = 1 To lngLastRow
        If Rows(lngRow).Hidden = True Then
            Set rngHidden = Union(rngHidden, Rows(lngRow))
        End If
    Next lngRow
    
    'Unhide everything in the hidden range variable.
    rngHidden.EntireRow.Hidden = False
    
    'Perform the sort on all the data.
    Range("A1").CurrentRegion.Sort _
        key1:=Range("A2"), _
        order1:=xlAscending, _
        header:=xlYes
        
    'Re-hide the rows that were originally hidden, but unhide the header.
    rngHidden.EntireRow.Hidden = True
    Rows(1).Hidden = False
    
    'Turn screen updating back on.
    Set rngHidden = Nothing
    Application.ScreenUpdating = True
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Range Object](range-object-excel.md)

