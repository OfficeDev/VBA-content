---
title: Page.Rectangles Property (Word)
keywords: vbawd10.chm11075590
f1_keywords:
- vbawd10.chm11075590
ms.prod: word
api_name:
- Word.Page.Rectangles
ms.assetid: 57c2f9f9-b858-b2f7-2dcc-1cbd565d009c
ms.date: 06/08/2017
---


# Page.Rectangles Property (Word)

Returns a  **Rectangles** collection that represents a portion of text or graphics in a page in a document.


## Syntax

 _expression_ . **Rectangles**

 _expression_ Required. A variable that represents a **[Page](page-object-word.md)** object.


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


[Page Object](page-object-word.md)

