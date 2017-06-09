---
title: Font.Attributes Property (Visio)
keywords: vis_sdr.chm12013100
f1_keywords:
- vis_sdr.chm12013100
ms.prod: visio
api_name:
- Visio.Font.Attributes
ms.assetid: 4d94e0d3-85a6-369f-5e04-83c9681c43c4
ms.date: 06/08/2017
---


# Font.Attributes Property (Visio)

Returns the attributes of the a  **Font** object. Read-only.


## Syntax

 _expression_ . **Attributes**

 _expression_ A variable that represents a **Font** object.


### Return Value

Integer


## Remarks

When you get the  **Attributes** property of a **Font** object, the following value is returned.



|**Constant**|**Value**|
|:-----|:-----|
| **visFont0Alias**|128|
A font marked as the font 0 alias is used instead of font 0 (the default font). The font 0 alias is used in some localized versions of Microsoft Visio and is controlled by means of entries in the registry.


