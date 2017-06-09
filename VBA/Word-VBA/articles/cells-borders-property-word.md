---
title: Cells.Borders Property (Word)
keywords: vbawd10.chm155845708
f1_keywords:
- vbawd10.chm155845708
ms.prod: word
api_name:
- Word.Cells.Borders
ms.assetid: df873357-9474-8f69-ae71-6df5859cbf93
ms.date: 06/08/2017
---


# Cells.Borders Property (Word)

Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified object.


## Syntax

 _expression_ . **Borders**

 _expression_ A variable that represents a **[Cells](cells-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example applies inside and outside borders to the cells in the first row of the first table in the active document.


```vb
Dim objTable As Table 
Set objTable = ActiveDocument.Tables(1) 
With objTable.Rows(1).Cells.Borders 
 .InsideLineStyle = wdLineStyleSingle 
 .OutsideLineStyle = wdLineStyleDouble 
End With
```


## See also


#### Concepts


[Cells Collection Object](cells-object-word.md)

