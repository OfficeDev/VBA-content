---
title: Frameset.ChildFramesetCount Property (Word)
keywords: vbawd10.chm165806085
f1_keywords:
- vbawd10.chm165806085
ms.prod: word
api_name:
- Word.Frameset.ChildFramesetCount
ms.assetid: 2e6bc910-9159-d3db-a399-0abc6bd9ba20
ms.date: 06/08/2017
---


# Frameset.ChildFramesetCount Property (Word)

Returns the number of child  **Frameset** objects associated with the specified **Frameset** object. Read-only **Long** .


## Syntax

 _expression_ . **ChildFramesetCount**

 _expression_ A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Remarks

This property applies only to  **Frameset** objects of type **wdFramesetTypeFrameset** . For more information on creating frames pages, see[Creating Frames Pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example displays the number of child Frameset objects contained by the Frameset object that represents the specified frames page.


```vb
MsgBox ActiveWindow.Document_ 
 .Frameset.ChildFramesetCount
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

