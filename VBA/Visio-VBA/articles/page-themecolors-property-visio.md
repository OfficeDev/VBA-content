---
title: Page.ThemeColors Property (Visio)
keywords: vis_sdr.chm10960180
f1_keywords:
- vis_sdr.chm10960180
ms.prod: visio
api_name:
- Visio.Page.ThemeColors
ms.assetid: a3f4bc4e-3dbb-9d50-9d71-f77b39ec0ac3
ms.date: 06/08/2017
---


# Page.ThemeColors Property (Visio)

Gets or sets the current theme colors for the page. Read/write.


## Syntax

 _expression_ . **ThemeColors**

 _expression_ An expression that returns a **Page** object.


### Return Value

Variant


## Remarks

You can set the  **ThemeColors** property value to any one of the following:




- The name or universal name of the theme color (strings)
    
- An enumerated value from the  **[VisThemeColors](visthemecolors-enumeration-visio.md)** enumeration
    
- A  **Master** object of type **visTypeThemeColors**
    


The  **ThemeColors** property always returns the universal name of the current theme colors.


