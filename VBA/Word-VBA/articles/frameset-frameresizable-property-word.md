---
title: Frameset.FrameResizable Property (Word)
keywords: vbawd10.chm165806111
f1_keywords:
- vbawd10.chm165806111
ms.prod: word
api_name:
- Word.Frameset.FrameResizable
ms.assetid: 5a373e57-3193-c2a3-52b6-42702237f6c3
ms.date: 06/08/2017
---


# Frameset.FrameResizable Property (Word)

 **True** if the user can resize the specified frame when the frames page is viewed in a Web browser. Read/write **Boolean** .


## Syntax

 _expression_ . **FrameResizable**

 _expression_ A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example sets the specified frame to be resizable when viewed in a Web browser.


```vb
With ActiveDocument.ActiveWindow.ActivePane.Frameset 
 .FrameDefaultURL = "C:\Documents\Order.htm" 
 .FrameResizable = True 
End With
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

