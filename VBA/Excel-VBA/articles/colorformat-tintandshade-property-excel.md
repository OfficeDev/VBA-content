---
title: ColorFormat.TintAndShade Property (Excel)
keywords: vbaxl10.chm105006
f1_keywords:
- vbaxl10.chm105006
ms.prod: excel
api_name:
- Excel.ColorFormat.TintAndShade
ms.assetid: b548b2ad-da3d-0d02-249e-2ab37271a5c6
ms.date: 06/08/2017
---


# ColorFormat.TintAndShade Property (Excel)

Returns or sets a  **Single** that lightens or darkens a color.


## Syntax

 _expression_ . **TintAndShade**

 _expression_ A variable that represents a **ColorFormat** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the  **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1 results in a run-time error: "The specified value is out of range." This property works for both theme colors and nontheme colors.


## See also


#### Concepts


[ColorFormat Object](colorformat-object-excel.md)

