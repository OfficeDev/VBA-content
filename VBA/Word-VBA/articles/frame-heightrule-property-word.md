---
title: Frame.HeightRule Property (Word)
keywords: vbawd10.chm153747457
f1_keywords:
- vbawd10.chm153747457
ms.prod: word
api_name:
- Word.Frame.HeightRule
ms.assetid: f7b96439-6e08-ee9c-3c77-739666756c50
ms.date: 06/08/2017
---


# Frame.HeightRule Property (Word)

Returns or sets a  **WdFrameSizeRule** that represents the rule for determining the height of the specified frame. Read/write.


## Syntax

 _expression_ . **HeightRule**

 _expression_ Required. A variable that represents a **[Frame](frame-object-word.md)** object.


## Example

This example sets both the height and width of the first frame in the active document to exactly 1 inch.


```vb
If ActiveDocument.Frames.Count >= 1 Then 
 With ActiveDocument.Frames(1) 
 .HeightRule = wdFrameExact 
 .Height = InchesToPoints(1) 
 .WidthRule = wdFrameExact 
 .Width = InchesToPoints(1) 
 End With 
End If
```


## See also


#### Concepts


[Frame Object](frame-object-word.md)

