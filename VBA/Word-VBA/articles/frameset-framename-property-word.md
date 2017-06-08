---
title: Frameset.FrameName Property (Word)
keywords: vbawd10.chm165806114
f1_keywords:
- vbawd10.chm165806114
ms.prod: word
api_name:
- Word.Frameset.FrameName
ms.assetid: f0b22dfe-3d12-0f75-1af2-23467b83a4ad
ms.date: 06/08/2017
---


# Frameset.FrameName Property (Word)

Returns or sets the name of the specified frame on a frames page. Read/write  **String** .


## Syntax

 _expression_ . **FrameName**

 _expression_ A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example sets the name of the specified frame to "BottomFrame".


```vb
ActiveWindow.Document.Frameset _ 
 .ChildFramesetItem(3).FrameName = "BottomFrame"
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

