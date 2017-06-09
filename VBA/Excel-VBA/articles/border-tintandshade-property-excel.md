---
title: Border.TintAndShade Property (Excel)
keywords: vbaxl10.chm547078
f1_keywords:
- vbaxl10.chm547078
ms.prod: excel
api_name:
- Excel.Border.TintAndShade
ms.assetid: 3ec15506-3ba6-a173-a11b-d17448fcdb1b
ms.date: 06/08/2017
---


# Border.TintAndShade Property (Excel)

Returns or sets a  **Single** that lightens or darkens a color.


## Syntax

 _expression_ . **TintAndShade**

 _expression_ A variable that represents a **Border** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the  **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1 results in a run-time error: "The specified value is out of range." This property works for both theme colors and nontheme colors.


## See also


#### Concepts


[Border Object](border-object-excel.md)

