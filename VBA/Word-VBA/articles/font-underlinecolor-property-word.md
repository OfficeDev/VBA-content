---
title: Font.UnderlineColor Property (Word)
keywords: vbawd10.chm156369062
f1_keywords:
- vbawd10.chm156369062
ms.prod: word
api_name:
- Word.Font.UnderlineColor
ms.assetid: f0da061c-0948-1214-ecdc-80f9c482f468
ms.date: 06/08/2017
---


# Font.UnderlineColor Property (Word)

Returns or sets the 24-bit color of the underline for the specified  **Font** object. .


## Syntax

 _expression_ . **UnderlineColor**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function. Setting the **UnderlineColor** property to **wdColorAutomatic** resets the color of the underline to the color of the text above it.


## See also


#### Concepts


[Font Object](font-object-word.md)

