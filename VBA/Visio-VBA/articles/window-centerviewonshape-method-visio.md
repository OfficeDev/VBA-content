---
title: Window.CenterViewOnShape Method (Visio)
keywords: vis_sdr.chm11662275
f1_keywords:
- vis_sdr.chm11662275
ms.prod: visio
api_name:
- Visio.Window.CenterViewOnShape
ms.assetid: 23f219be-bfb7-0f5b-89c0-855093e4bbd9
ms.date: 06/08/2017
---


# Window.CenterViewOnShape Method (Visio)

Pans the Microsoft Visio drawing window to place the specified shape in the center of the view.


## Syntax

 _expression_ . **CenterViewOnShape**( **_SheetObject_** , **_Flags_** )

 _expression_ A variable that represents a **[Window](window-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SheetObject_|Required| **[Shape](shape-object-visio.md)**|The shape to center in the view.|
| _Flags_|Required| **[VisCenterViewFlags](viscenterviewflags-enumeration-visio.md)**|The centering behavior to apply.|

### Return Value

 **Nothing**


## Remarks

The  _Flags_ parameter value must be combination of one of more of the following **VisCenterViewFlags** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visCenterViewDefault**|0|Display the page that contains the specified shape and center the view on the shape.|
| **visCenterViewIfOffScreen**|1|Center the view only if the shape is currently off screen.|
| **visCenterViewSelectShape**|2|Also select the shape.|
If the specified shape is not valid, Microsoft Visio returns an Invalid Parameter error.


