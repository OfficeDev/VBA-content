---
title: Table.Columns Property (Word)
keywords: vbawd10.chm156303460
f1_keywords:
- vbawd10.chm156303460
ms.prod: word
api_name:
- Word.Table.Columns
ms.assetid: 6f4c70ef-032d-7f05-1b21-c5c86af804bd
ms.date: 06/08/2017
---


# Table.Columns Property (Word)

Returns a  **[Columns](columns-object-word.md)** collection that represents all the table columns in the table. Read-only.


## Syntax

 _expression_ . **Columns**

 _expression_ A variable that represents a **[Table](table-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the number of columns in the first table in the active document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 MsgBox ActiveDocument.Tables(1).Columns.Count 
End If
```


## See also


#### Concepts


[Table Object](table-object-word.md)

