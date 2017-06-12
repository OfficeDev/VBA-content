---
title: Frame.WidthRule Property (Word)
keywords: vbawd10.chm153747458
f1_keywords:
- vbawd10.chm153747458
ms.prod: word
api_name:
- Word.Frame.WidthRule
ms.assetid: cd780bff-f0b9-c594-a134-005f3cce2edf
ms.date: 06/08/2017
---


# Frame.WidthRule Property (Word)

Returns or sets the rule used to determine the width of a frame. Read/write  **WdFrameSizeRule** .


## Syntax

 _expression_ . **WidthRule**

 _expression_ Required. A variable that represents a **[Frame](frame-object-word.md)** object.


## Example

This example sets the width of the last frame in the active document to exactly 72 points (1 inch).


```vb
If ActiveDocument.Frames.Count >= 1 Then 
 With ActiveDocument.Frames(ActiveDocument.Frames.Count) 
 .WidthRule = wdFrameExact 
 .Width = 72 
 End With 
End If
```


## See also


#### Concepts


[Frame Object](frame-object-word.md)

