---
title: Frameset.FrameDefaultURL Property (Word)
keywords: vbawd10.chm165806116
f1_keywords:
- vbawd10.chm165806116
ms.prod: word
api_name:
- Word.Frameset.FrameDefaultURL
ms.assetid: 596f57d4-2514-8cd0-2d97-20618051fd6c
ms.date: 06/08/2017
---


# Frameset.FrameDefaultURL Property (Word)

Returns or sets the Web page or other document to be displayed in the specified frame when the frames page is opened. Read/write  **String** .


## Syntax

 _expression_ . **FrameDefaultURL**

 _expression_ A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example sets the specified frame to display a local file named "Order.htm".


```vb
With ActiveDocument.ActiveWindow.ActivePane.Frameset 
 .FrameDefaultURL = "C:\Documents\Order.htm" 
 .FrameLinkToFile = True 
End With
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

