---
title: Frameset.FrameLinkToFile Property (Word)
keywords: vbawd10.chm165806117
f1_keywords:
- vbawd10.chm165806117
ms.prod: word
api_name:
- Word.Frameset.FrameLinkToFile
ms.assetid: a27ce637-a892-3697-a727-e7c60eb26aaf
ms.date: 06/08/2017
---


# Frameset.FrameLinkToFile Property (Word)

 **True** if the Web page or other document specified by the **[FrameDefaultURL](frameset-framedefaulturl-property-word.md)** property is an external file to which Microsoft Word maintains only a link from the specified frame. Read/write **Boolean** .


## Syntax

 _expression_ . **FrameLinkToFile**

 _expression_ A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example sets Microsoft Word to maintain only a link from the specified frame to the document "Order.htm".


```vb
With ActiveDocument.ActiveWindow.ActivePane.Frameset 
 .FrameDefaultURL = "C:\Documents\Order.htm" 
 .FrameLinkToFile = True 
End With
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

