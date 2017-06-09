---
title: Page.PasteToLocation Method (Visio)
keywords: vis_sdr.chm10962120
f1_keywords:
- vis_sdr.chm10962120
ms.prod: visio
api_name:
- Visio.Page.PasteToLocation
ms.assetid: d24cc1b3-c0c7-d529-b94f-0fea82d124ef
ms.date: 06/08/2017
---


# Page.PasteToLocation Method (Visio)

Pastes the shape to the specified location on the page.


## Syntax

 _expression_ . **PasteToLocation**( **_xPos_** , **_yPos_** , **_Flags_** )

 _expression_ A variable that represents a **[Page](page-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _xPos_|Required| **Double**|The x-coordinate at which to place the center of the object?s width or PinX, in inches.|
| _yPos_|Required| **Double**|The y-coordinate at which to place the center of the object?s height or PinY, in inches.|
| _Flags_|Required| **Long**|The default is zero.|

### Return Value

 **Nothing**


## Remarks

The  _Flags_ parameter value can also be **visCopyPasteDontAddToContainers** .


