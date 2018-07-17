---
title: InvisibleApp.GetBuiltInStencilFile Method (Visio)
keywords: vis_sdr.chm17562110
f1_keywords:
- vis_sdr.chm17562110
ms.prod: visio
api_name:
- Visio.InvisibleApp.GetBuiltInStencilFile
ms.assetid: 2f8e28a9-67bd-31fd-25f1-f684dfeeeca8
ms.date: 06/08/2017
---


# InvisibleApp.GetBuiltInStencilFile Method (Visio)

Returns the file path to the specified built-in, hidden stencil used to populate certain galleries in the Microsoft Visio user interface.


## Syntax

 _expression_ . **GetBuiltInStencilFile**( **_StencilType_** , **_MeasurementSystem_** )

 _expression_ A variable that represents an **[InvisibleApp](invisibleapp-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StencilType_|Required| **[VisBuiltInStencilTypes](visbuiltinstenciltypes-enumeration-visio.md)**|The stencil to retrieve. See Remarks for possible values.|
| _MeasurementSystem_|Required| **[VisMeasurementSystem](vismeasurementsystem-enumeration-visio.md)**|The measurement system for the stencil.|

### Return Value

 **String**


## Remarks

The  _StencilType_ parameter value must be one of the following **VisBuiltInStencilTypes** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visBuiltInStencilBackgrounds**|0|The hidden stencil that contains the shapes displayed in the  **Backgrounds** gallery ( **Design** tab).|
| **visBuiltInStencilBorders**|1|The hidden stencil that contains the shapes displayed in the  **Borders and Titles** gallery ( **Design** tab).|
| **visBuiltInStencilContainers**|2|The hidden stencil that contains the shapes displayed in the  **Container** gallery ( **Insert** tab).|
| **visBuiltInStencilCallouts**|3|The hidden stencil that contains the shapes displayed in the  **Callout** gallery ( **Insert** tab).|
| **visBuiltInStencilLegends**|4|The hidden stencil that contains the shapes displayed in the  **Insert Legend** gallery ( **Data** tab).|

