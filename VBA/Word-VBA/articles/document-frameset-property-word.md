---
title: Document.Frameset Property (Word)
keywords: vbawd10.chm158007623
f1_keywords:
- vbawd10.chm158007623
ms.prod: word
api_name:
- Word.Document.Frameset
ms.assetid: 40079f4f-be1d-c8dd-5536-ccb5f570bde9
ms.date: 06/08/2017
---


# Document.Frameset Property (Word)

Returns a  **[Frameset](frameset-object-word.md)** object that represents an entire frames page or a single frame on a frames page. Read-only.


## Syntax

 _expression_ . **Frameset**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating frames pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example sets the color of frame borders in the specified frames page to tan.


```vb
With ActiveWindow.Document.Frameset 
 .FramesetBorderColor = wdColorTan 
 .FramesetBorderWidth = 6 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

