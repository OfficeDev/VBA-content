---
title: Pane.Frameset Property (Word)
keywords: vbawd10.chm157286418
f1_keywords:
- vbawd10.chm157286418
ms.prod: word
api_name:
- Word.Pane.Frameset
ms.assetid: 6bab63ae-aa83-e2b8-9b92-e472c2433246
ms.date: 06/08/2017
---


# Pane.Frameset Property (Word)

Returns a  **[Frameset](frameset-object-word.md)** object that represents an entire frames page or a single frame on a frames page. Read-only.


## Syntax

 _expression_ . **Frameset**

 _expression_ A variable that represents a **[Pane](pane-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating frames pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).




## Example

This example adds a new frame to the immediate right of the specified frame.


```vb
ActiveDocument.ActiveWindow.ActivePane.Frameset _ 
 .AddNewFrame wdFramesetNewRight
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

