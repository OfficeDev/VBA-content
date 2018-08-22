---
title: Worksheet.Rows Property (Excel)
keywords: vbaxl10.chm175122
f1_keywords:
- vbaxl10.chm175122
ms.prod: excel
api_name:
- Excel.Worksheet.Rows
ms.assetid: 5d07304e-a3c9-2a75-b2ba-4a7b16ce6516
ms.date: 06/08/2017
---


# Worksheet.Rows Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents all the rows on the specified worksheet.


## Syntax

 _expression_ . **Rows**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

`Rows` without an object qualifier is equivalent to `ActiveSheet.Rows`. If the active sheet isn't a worksheet, then `Rows` fails.

To return a single row, include an index in parentheses. For example, `Rows(1)` returns the first row.


## Example

This example deletes row three on Sheet1.


```vb
Worksheets("Sheet1").Rows(3).Delete
```

This example deletes rows in the current region on worksheet one where the value of cell one in the row is the same as the value in cell one in the previous row.




```vb
For Each rw In Worksheets(1).Cells(1, 1).CurrentRegion.Rows 
 this = rw.Cells(1, 1).Value 
 If this = last Then rw.Delete 
 last = this 
Next
```

This example displays the number of rows in the selection on Sheet1. If more than one area is selected, the example loops through each area.




```vb
Worksheets("Sheet1").Activate 
areaCount = Selection.Areas.Count 
If areaCount <= 1 Then 
 MsgBox "The selection contains " &; _ 
 Selection.Rows.Count &; " rows." 
Else 
 i = 1 
 For Each a In Selection.Areas 
 MsgBox "Area " &; i &; " of the selection contains " &; _ 
 a.Rows.Count &; " rows." 
 i = i + 1 
 Next a 
End If
```


## See also


#### Related

[Range.Rows Property](range-rows-property-excel.md)


#### Concepts

[Worksheet Object](worksheet-object-excel.md)

