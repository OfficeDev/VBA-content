---
title: TextFrame.WarpFormat Property (Word)
keywords: vbawd10.chm162665366
f1_keywords:
- vbawd10.chm162665366
ms.prod: word
api_name:
- Word.TextFrame.WarpFormat
ms.assetid: 2ea707b9-0ed1-1196-2bf9-a32ae87d456a
ms.date: 06/08/2017
---


# TextFrame.WarpFormat Property (Word)

Returns or sets the warp format (how the text is warped) for the specified text frame. Read/write [MsoWarpFormat](http://msdn.microsoft.com/library/481cead3-900f-66b6-8200-21342b0ce21c%28Office.15%29.aspx).


## Syntax

 _expression_ . **WarpFormat**

 _expression_ A variable that represents a **[TextFrame](textframe-object-word.md)** object.


## Example

The following code example shows how to set the warp format for the first shape on the active document.


```vb
ActiveDocument.Shapes(1).TextFrame.WarpFormat = msoWarpFormat15
```


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

