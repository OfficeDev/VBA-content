---
title: Range.Rows Property (Excel)
keywords: vbaxl10.chm144191
f1_keywords:
- vbaxl10.chm144191
ms.prod: excel
api_name:
- Excel.Range.Rows
ms.assetid: 2b0541f1-119d-8535-8418-ff9482353ec1
ms.date: 06/08/2017
---


# Range.Rows Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the rows in the specified range. Read-only **Range** object.


## Syntax

 _expression_ . **Rows**

 _expression_ A variable that represents a **Range** object.


## Remarks

Using this property without an object qualifier is equivalent to using  `ActiveSheet.Rows`.

When applied to a  **Range** object that's a multiple selection, this property returns rows from only the first area of the range. For example, if the **Range** object has two areas — A1:B2 and C3:D4 — `Selection.Rows.Count` returns 2, not 4. To use this property on a range that may contain a multiple selection, test `Areas.Count` to determine whether the range is a multiple selection. If it is, loop over each area in the range, as shown in the third example.


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


#### Concepts


[Range Object](range-object-excel.md)

