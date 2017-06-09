---
title: Frameset.Width Property (Word)
keywords: vbawd10.chm165806083
f1_keywords:
- vbawd10.chm165806083
ms.prod: word
api_name:
- Word.Frameset.Width
ms.assetid: 08c2c81a-119f-18ab-fa6e-5a21ab673cba
ms.date: 06/08/2017
---


# Frameset.Width Property (Word)

Returns or sets the width (in points) of the specified  **Frameset** object. Read/write **Long** .


## Syntax

 _expression_ . **Width**

 _expression_ A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Remarks

Use the  **[WidthType](frameset-widthtype-property-word.md)** property to specify the type of unit in which this value is expressed.


## Example

This example sets the width of the specified  **Frameset** object to 25% of the window width.


```vb
With ActiveWindow.ActivePane.Frameset 
 .WidthType = wdFramesetSizeTypePercent 
 .Width = 25 
End With
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

