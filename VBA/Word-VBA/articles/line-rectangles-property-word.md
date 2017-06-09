---
title: Line.Rectangles Property (Word)
keywords: vbawd10.chm49610760
f1_keywords:
- vbawd10.chm49610760
ms.prod: word
api_name:
- Word.Line.Rectangles
ms.assetid: 2cc7988f-de81-5461-5fd0-a0ce29fdadaa
ms.date: 06/08/2017
---


# Line.Rectangles Property (Word)

Returns a  **Rectangles** collection that represents a portion of text or graphics in a page in a document.


## Syntax

 _expression_ . **Rectangles**

 _expression_ Required. A variable that represents a **[Line](line-object-word.md)** object.


## Remarks

Use the  **Rectangles** collection and related objects and properties for programmatically defining page layout in a document. Rectangles correspond to portions of text or graphics on a page in a document.


## Example

The following example returns the  **Rectangles** collection for the first page in the active document.


```vb
Dim objRectangles As Rectangles 
 
Set objRectangles = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles
```


## See also


#### Concepts


[Line Object](line-object-word.md)

