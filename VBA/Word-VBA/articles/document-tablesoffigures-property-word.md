---
title: Document.TablesOfFigures Property (Word)
keywords: vbawd10.chm158007321
f1_keywords:
- vbawd10.chm158007321
ms.prod: word
api_name:
- Word.Document.TablesOfFigures
ms.assetid: 1c386611-82f9-0a0d-71ce-dfe006d8eab5
ms.date: 06/08/2017
---


# Document.TablesOfFigures Property (Word)

Returns a  **[TablesOfFigures](document-tablesoffigures-property-word.md)** collection that represents the tables of figures in the specified document. Read-only.


## Syntax

 _expression_ . **TablesOfFigures**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds a table of figures at the insertion point in the active document.


```
Selection.Collapse Direction:=wdCollapseStart 
ActiveDocument.TablesOfFigures.Add Range:=Selection.Range, _ 
 Caption:=wdCaptionFigure
```

This example updates the contents of the first table of figures in Report.doc.




```
Documents("Report.doc").TablesOfFigures(1).Update
```


## See also


#### Concepts


[Document Object](document-object-word.md)

