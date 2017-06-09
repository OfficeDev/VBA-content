---
title: Rows.Borders Property (Word)
keywords: vbawd10.chm155976780
f1_keywords:
- vbawd10.chm155976780
ms.prod: word
api_name:
- Word.Rows.Borders
ms.assetid: 4c251987-5bbb-bfdb-d90f-861838f1b59d
ms.date: 06/08/2017
---


# Rows.Borders Property (Word)

Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified object.


## Syntax

 _expression_ . **Borders**

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).




## Example

This example applies inside and outside borders to the rows in the first table in the active document.


```vb
Set myTable = ActiveDocument.Tables(1) 
With myTable.Rows.Borders 
 .InsideLineStyle = wdLineStyleSingle 
 .OutsideLineStyle = wdLineStyleDouble 
End With
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

