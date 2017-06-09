---
title: TableOfFigures.Caption Property (Word)
keywords: vbawd10.chm153157633
f1_keywords:
- vbawd10.chm153157633
ms.prod: word
api_name:
- Word.TableOfFigures.Caption
ms.assetid: 66848200-1eaa-f0ed-f270-51339de1f213
ms.date: 06/08/2017
---


# TableOfFigures.Caption Property (Word)

Returns or sets the label that identifies the items to be included in a table of figures. Read/write  **String** .


## Syntax

 _expression_ . **Caption**

 _expression_ A variable that represents a **[TableOfFigures](tableoffigures-object-word.md)** collection.


## Remarks

This property corresponds to the \c switch for a TOC field.


## Example

This example inserts a table caption and then changes the caption of the first table of figures to "Table."


```vb
Selection.Collapse Direction:=wdCollapseStart 
Selection.Range.InsertCaption "Table" 
If ActiveDocument.TablesOfFigures.Count >= 1 Then 
 ActiveDocument.TablesOfFigures(1).Caption = "Table" 
End If
```


## See also


#### Concepts


[TableOfFigures Object](tableoffigures-object-word.md)

