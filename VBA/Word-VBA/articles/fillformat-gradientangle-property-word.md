---
title: FillFormat.GradientAngle Property (Word)
keywords: vbawd10.chm164102267
f1_keywords:
- vbawd10.chm164102267
ms.prod: word
api_name:
- Word.FillFormat.GradientAngle
ms.assetid: a40a03bb-a395-0e7e-708c-4b9eee89ee4c
ms.date: 06/08/2017
---


# FillFormat.GradientAngle Property (Word)

Returns or sets the angle of the gradient fill for the specified fill format. Read/write.


## Syntax

 _expression_ . **GradientAngle**

 _expression_ An expression that returns a **[FillFormat](fillformat-object-word.md)** object.


### Return Value

Single


## Remarks

A gradient fill can be specified in the formatting for various elements (shapes) in a chart. For example, you can use the  **Format Data Series** dialog box to format the columns in a **Column** chart to a gradient fill. In this case, the **GradientAngle** property corresponds to the setting of the **Angle** box in the **Fill** category of the **Format Data Series** dialog box. The valid range of values for the **GradientAngle** property is from 0 to 359.9.


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

