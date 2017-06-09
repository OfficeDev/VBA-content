---
title: Page.GetTheme Method (Visio)
ms.prod: visio
ms.assetid: 31c84e69-0bc8-2d1a-84d8-7397110d74ae
ms.date: 06/08/2017
---


# Page.GetTheme Method (Visio)

Returns a  **Variant** that represents the specified theme component of the specified page.


## Syntax

 _expression_ . **GetTheme**_(eThemeType)_

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _eThemeType_|Required|[VISTHEMETYPES](visthemetypes-enumeration-visio.md)|Specifies the type of the theme component to return.|

### Return value

 **VARIANT**


## Remarks

The theme components returned are enumerations of built-in theme definitions for colors, fonts, and styles for 2-dimensional shapes, and styles for 1-dimensional shapes.


## See also


#### Concepts


[Page Object](page-object-visio.md)

