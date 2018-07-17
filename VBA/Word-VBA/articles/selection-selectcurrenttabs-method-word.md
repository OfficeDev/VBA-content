---
title: Selection.SelectCurrentTabs Method (Word)
keywords: vbawd10.chm158663177
f1_keywords:
- vbawd10.chm158663177
ms.prod: word
api_name:
- Word.Selection.SelectCurrentTabs
ms.assetid: 38b0ba64-eedc-9ef5-5622-5499b50bbd3e
ms.date: 06/08/2017
---


# Selection.SelectCurrentTabs Method (Word)

Extends the selection forward until a paragraph with different tab stops is encountered.


## Syntax

 _expression_ . **SelectCurrentTabs**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example selects the second paragraph in the active document and then extends the selection to include all other paragraphs that have the same tab stops.


```vb
Set myRange = ActiveDocument.Paragraphs(2).Range 
myRange.Select 
Selection.SelectCurrentTabs
```

This example selects paragraphs that have the same tab stops and retrieves the position of the first tab stop. The example moves the selection to the next range of paragraphs that have the same tab stops. The example then adds the tab stop setting from the first group of paragraphs to the current selection.




```vb
With Selection 
 .SelectCurrentTabs 
 pos = .Paragraphs.TabStops(1).Position 
 .Collapse Direction:=wdCollapseEnd 
 .SelectCurrentTabs 
 .Paragraphs.TabStops.Add Position:=pos 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

