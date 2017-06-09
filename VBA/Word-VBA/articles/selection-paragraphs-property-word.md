---
title: Selection.Paragraphs Property (Word)
keywords: vbawd10.chm158662715
f1_keywords:
- vbawd10.chm158662715
ms.prod: word
api_name:
- Word.Selection.Paragraphs
ms.assetid: f237788a-01e4-62ce-d698-3af619c90272
ms.date: 06/08/2017
---


# Selection.Paragraphs Property (Word)

Returns a  **[Paragraphs](paragraphs-object-word.md)** collection that represents all the paragraphs in the specified selection. Read-only.


## Syntax

 _expression_ . **Paragraphs**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the line spacing to double for the first paragraph in the selection.


```
Selection.Paragraphs(1).LineSpacingRule = wdLineSpaceDouble
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

