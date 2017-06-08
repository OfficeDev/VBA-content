---
title: Selection.Footnotes Property (Word)
keywords: vbawd10.chm158662710
f1_keywords:
- vbawd10.chm158662710
ms.prod: word
api_name:
- Word.Selection.Footnotes
ms.assetid: 61829c93-46e9-c1c5-1424-fb34a812a76d
ms.date: 06/08/2017
---


# Selection.Footnotes Property (Word)

Returns a  **[Footnotes](footnotes-object-word.md)** collection that represents all the footnotes in a range, selection, or document. Read-only.


## Syntax

 _expression_ . **Footnotes**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example inserts an automatically numbered footnote at the insertion point.


```
Selection.Collapse Direction:=wdCollapseStart 
Selection.Footnotes.Add Range:=Selection.Range, _ 
 Text:="(Lone Creek Press, 1995)"
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

