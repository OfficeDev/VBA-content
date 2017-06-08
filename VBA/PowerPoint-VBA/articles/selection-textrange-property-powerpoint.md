---
title: Selection.TextRange Property (PowerPoint)
keywords: vbapp10.chm508010
f1_keywords:
- vbapp10.chm508010
ms.prod: powerpoint
api_name:
- PowerPoint.Selection.TextRange
ms.assetid: 532c0a35-c18d-8030-8e6a-3f1cdb47c244
ms.date: 06/08/2017
---


# Selection.TextRange Property (PowerPoint)

Returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents the selected text. Read-only.


## Syntax

 _expression_. **TextRange**

 _expression_ A variable that represents a **Selection** object.


### Return Value

TextRange


## Remarks

You can construct a text range from a selection when the presentation is in slide view, normal view, outline view, notes page view, or any master view.


## Example

This example makes the selected text bold in the first window.


```
Windows(1).Selection.TextRange.Font.Bold = True
```


## See also


#### Concepts


[Selection Object](selection-object-powerpoint.md)

