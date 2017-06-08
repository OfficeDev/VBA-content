---
title: Selection.Rows Property (Word)
keywords: vbawd10.chm158662959
f1_keywords:
- vbawd10.chm158662959
ms.prod: word
api_name:
- Word.Selection.Rows
ms.assetid: 800edca7-fc0f-ed73-ae3a-400eadcccf8b
ms.date: 06/08/2017
---


# Selection.Rows Property (Word)

Returns a  **[Rows](rows-object-word.md)** collection that represents all the table rows in a range, selection, or table. Read-only.


## Syntax

 _expression_ . **Rows**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example places a border around the cells in the row that contains the insertion point.


```vb
Selection.Collapse Direction:=wdCollapseStart 
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows(1).Borders.OutsideLineStyle = wdLineStyleSingle 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

