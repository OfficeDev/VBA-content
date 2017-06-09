---
title: Range.Columns Property (Excel)
keywords: vbaxl10.chm144101
f1_keywords:
- vbaxl10.chm144101
ms.prod: excel
api_name:
- Excel.Range.Columns
ms.assetid: a1a23288-e911-909d-0bc0-48bdce2ccbac
ms.date: 06/08/2017
---


# Range.Columns Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the columns in the specified range.


## Syntax

 _expression_ . **Columns**

 _expression_ A variable that represents a **Range** object.


## Remarks

Using this property without an object qualifier is equivalent to using  `ActiveSheet.Columns`.

When applied to a  **Range** object that's a multiple-area selection, this property returns columns from only the first area of the range. For example, if the **Range** object has two areas — A1:B2 and C3:D4 — `Selection.Columns.Count` returns 2, not 4. To use this property on a range that may contain a multiple-area selection, test `Areas.Count` to determine whether the range contains more than one area. If it does, loop over each area in the range.


## Example

This example sets the value of every cell in column one in the range named "myRange" to 0 (zero).


```vb
Range("myRange").Columns(1).Value = 0
```

This example displays the number of columns in the selection on Sheet1. If more than one area is selected, the example loops through each area.




```vb
Worksheets("Sheet1").Activate 
areaCount = Selection.Areas.Count 
If areaCount <= 1 Then 
 MsgBox "The selection contains " &; _ 
 Selection.Columns.Count &; " columns." 
Else 
 For i = 1 To areaCount 
 MsgBox "Area " &; i &; " of the selection contains " &; _ 
 Selection.Areas(i).Columns.Count &; " columns." 
 Next i 
End If
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

