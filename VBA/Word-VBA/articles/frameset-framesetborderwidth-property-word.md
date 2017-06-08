---
title: Frameset.FramesetBorderWidth Property (Word)
keywords: vbawd10.chm165806100
f1_keywords:
- vbawd10.chm165806100
ms.prod: word
api_name:
- Word.Frameset.FramesetBorderWidth
ms.assetid: 6d828372-78a3-83f1-d162-91b000af2023
ms.date: 06/08/2017
---


# Frameset.FramesetBorderWidth Property (Word)

Returns or sets the width (in points) of the borders surrounding the frames on the specified frames page. Read/write  **Single** .


## Syntax

 _expression_ . **FramesetBorderWidth**

 _expression_ A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example sets the width of frame borders in the specified frames page to 6 points.


```vb
With ActiveWindow.Document.Frameset 
 .FramesetBorderColor = wdColorTan 
 .FramesetBorderWidth = 6 
End With
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

