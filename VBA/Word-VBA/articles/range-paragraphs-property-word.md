---
title: Range.Paragraphs Property (Word)
keywords: vbawd10.chm157155387
f1_keywords:
- vbawd10.chm157155387
ms.prod: word
api_name:
- Word.Range.Paragraphs
ms.assetid: b5c9df62-a477-ce1a-4a94-027100527a6f
ms.date: 06/08/2017
---


# Range.Paragraphs Property (Word)

Returns a  **Paragraphs** collection that represents all the paragraphs in the specified range. Read-only.


## Syntax

 _expression_ . **Paragraphs**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the line spacing to single for the collection of all paragraphs in section one in the active document.


```vb
ActiveDocument.Sections(1).Range.Paragraphs.LineSpacingRule = _ 
 wdLineSpaceSingle
```


## See also


#### Concepts


[Range Object](range-object-word.md)

