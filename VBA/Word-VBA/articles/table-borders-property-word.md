---
title: Table.Borders Property (Word)
keywords: vbawd10.chm156304460
f1_keywords:
- vbawd10.chm156304460
ms.prod: word
api_name:
- Word.Table.Borders
ms.assetid: 904bce6b-db91-32be-f65d-7200f9a63be8
ms.date: 06/08/2017
---


# Table.Borders Property (Word)

Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified object.


## Syntax

 _expression_ . **Borders**

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example applies inside and outside borders to the first table in the active document.


```vb
Set myTable = ActiveDocument.Tables(1) 
With myTable.Borders 
 .InsideLineStyle = wdLineStyleSingle 
 .OutsideLineStyle = wdLineStyleDouble 
End With
```


## See also


#### Concepts


[Table Object](table-object-word.md)

