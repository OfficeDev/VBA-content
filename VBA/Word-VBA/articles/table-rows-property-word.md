---
title: Table.Rows Property (Word)
keywords: vbawd10.chm156303461
f1_keywords:
- vbawd10.chm156303461
ms.prod: word
api_name:
- Word.Table.Rows
ms.assetid: e4cc7541-15fe-97b6-0fe6-90d561a85420
ms.date: 06/08/2017
---


# Table.Rows Property (Word)

Returns a  **Rows** collection that represents all the table rows within a table. Read-only.


## Syntax

 _expression_ . **Rows**

 _expression_ A variable that represents a **[Table](table-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example deletes the second row from the first table in the active document.


```vb
ActiveDocument.Tables(1).Rows(2).Delete
```


## See also


#### Concepts


[Table Object](table-object-word.md)

