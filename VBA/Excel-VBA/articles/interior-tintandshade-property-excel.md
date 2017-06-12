---
title: Interior.TintAndShade Property (Excel)
keywords: vbaxl10.chm551080
f1_keywords:
- vbaxl10.chm551080
ms.prod: excel
api_name:
- Excel.Interior.TintAndShade
ms.assetid: 45b12e93-1a6d-b5a3-b31d-4b41d87f3f73
ms.date: 06/08/2017
---


# Interior.TintAndShade Property (Excel)

Returns or sets a  **Single** that lightens or darkens a color.


## Syntax

 _expression_ . **TintAndShade**

 _expression_ A variable that represents an **Interior** object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the  **TintAndShade** property. Zero (0) is neutral.

Attempting to set this property to a value less than -1 or more than 1 results in a run-time error: "The specified value is out of range." This property works for both theme colors and nontheme colors.


## See also


#### Concepts


[Interior Object](interior-object-excel.md)

