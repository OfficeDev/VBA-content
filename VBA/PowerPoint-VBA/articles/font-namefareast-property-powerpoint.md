---
title: Font.NameFarEast Property (PowerPoint)
keywords: vbapp10.chm575016
f1_keywords:
- vbapp10.chm575016
ms.prod: powerpoint
api_name:
- PowerPoint.Font.NameFarEast
ms.assetid: 0b3f7d98-bda5-eec3-f570-20d8b575c0a3
ms.date: 06/08/2017
---


# Font.NameFarEast Property (PowerPoint)

Returns or sets the Asian font name. Read/write.


## Syntax

 _expression_. **NameFarEast**

 _expression_ A variable that represents a **Font** object.


### Return Value

String


## Remarks

Use the  **[Replace](fonts-replace-method-powerpoint.md)** method to change the font that's applied to all text and that appears in the **Font** box on the **Font** tab.


## Example

This example displays the name of the Asian font applied to the selection.


```vb
MsgBox ActiveWindow.Selection.ShapeRange _
    .TextFrame.TextRange.Font.NameFarEast
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

