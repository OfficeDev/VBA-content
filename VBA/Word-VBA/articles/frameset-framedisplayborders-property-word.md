---
title: Frameset.FrameDisplayBorders Property (Word)
keywords: vbawd10.chm165806115
f1_keywords:
- vbawd10.chm165806115
ms.prod: word
api_name:
- Word.Frameset.FrameDisplayBorders
ms.assetid: a1993b72-2737-92d8-d1bc-b4bc0182b23a
ms.date: 06/08/2017
---


# Frameset.FrameDisplayBorders Property (Word)

 **True** if the frame borders on the specified frames page are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **FrameDisplayBorders**

 _expression_ A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example sets Microsoft Word to display frame borders in the specified frames page.


```vb
ActiveDocument.ActiveWindow.ActivePane.Frameset _ 
 .FrameDisplayBorders = True
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

