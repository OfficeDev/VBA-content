---
title: Range.ListHeaderRows Property (Excel)
keywords: vbaxl10.chm144154
f1_keywords:
- vbaxl10.chm144154
ms.prod: excel
api_name:
- Excel.Range.ListHeaderRows
ms.assetid: d71a9b28-cd5d-677c-9ce1-f8de2b350e5f
ms.date: 06/08/2017
---


# Range.ListHeaderRows Property (Excel)

Returns the number of header rows for the specified range. Read-only  **Long** .


## Syntax

 _expression_ . **ListHeaderRows**

 _expression_ A variable that represents a **Range** object.


## Remarks

Before you use this property, use the  **[CurrentRegion](range-currentregion-property-excel.md)** property to find the boundaries of the range.


## Example

This example sets the  `rTbl` variable to the range represented by the current region for the active cell, not including any header rows.


```vb
Set rTbl = ActiveCell.CurrentRegion 
' remove the headers from the range 
iHdrRows = rTbl.ListHeaderRows 
If iHdrRows > 0 Then 
 ' resize the range minus n rows 
 Set rTbl = rTbl.Resize(rTbl.Rows.Count - iHdrRows) 
 ' and then move the resized range down to 
 ' get to the first non-header row 
 Set rTbl = rTbl.Offset(iHdrRows) 
End If
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

