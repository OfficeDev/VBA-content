---
title: Page.SetThemeVariant Method (Visio)
ms.prod: visio
ms.assetid: 8393a95f-83ca-0efa-d987-ae498bfe5e9d
ms.date: 06/08/2017
---


# Page.SetThemeVariant Method (Visio)

Sets the color, style, and optionally the embellishment of the variant of the theme applied to the specified page.


## Syntax

 _expression_ . **SetThemeVariant**_(variantColor,_ _variantStyle,_ _embellishment)_

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _variantColor_|Required|INT16|The index of the color variant to apply. Possible values are from 0 to 3.|
| _variantStyle_|Required|INT16|The index of the style variant to apply. Possible values are from 0 to 3.|
| _embellishment_|Optional|INT16|The index of the embellishment to apply. Possible values are from 1, for ?low,? to 3, for ?high.?|

### Return value

 **VOID**


## See also


#### Concepts


[Page Object](page-object-visio.md)

