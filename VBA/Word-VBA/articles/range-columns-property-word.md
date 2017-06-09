---
title: Range.Columns Property (Word)
keywords: vbawd10.chm157155630
f1_keywords:
- vbawd10.chm157155630
ms.prod: word
api_name:
- Word.Range.Columns
ms.assetid: 667b808a-e885-a7b7-0a68-5b2466ddd869
ms.date: 06/08/2017
---


# Range.Columns Property (Word)

Returns a  **[Columns](columns-object-word.md)** collection that represents all the table columns in the range. Read-only.


## Syntax

 _expression_ . **Columns**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the number of columns in the first table in the active document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 MsgBox ActiveDocument.Tables(1).Columns.Count 
End If
```

This example sets the width of the current column to 1 inch.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Columns.SetWidth ColumnWidth:=InchesToPoints(1), _ 
 RulerStyle:=wdAdjustProportional 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

