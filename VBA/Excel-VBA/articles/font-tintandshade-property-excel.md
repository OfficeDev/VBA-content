---
title: Font.TintAndShade Property (Excel)
keywords: vbaxl10.chm559088
f1_keywords:
- vbaxl10.chm559088
ms.prod: excel
api_name:
- Excel.Font.TintAndShade
ms.assetid: 0b890357-fb55-ac43-ecf0-f7d84ce0f248
ms.date: 06/08/2017
---


# Font.TintAndShade Property (Excel)

Returns or sets a  **Single** that lightens or darkens a color.


## Syntax

 _expression_ . **TintAndShade**

 _expression_ A variable that represents a **Font** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the  **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1 results in a run-time error: "The specified value is out of range." This property works for both theme colors and nontheme colors.


## See also


#### Concepts


[Font Object](font-object-excel.md)

