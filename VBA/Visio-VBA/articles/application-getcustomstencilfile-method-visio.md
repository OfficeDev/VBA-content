---
title: Application.GetCustomStencilFile Method (Visio)
keywords: vis_sdr.chm10062115
f1_keywords:
- vis_sdr.chm10062115
ms.prod: visio
api_name:
- Visio.Application.GetCustomStencilFile
ms.assetid: 10c8ec1d-f4e0-07dd-4487-40f85cbf5497
ms.date: 06/08/2017
---


# Application.GetCustomStencilFile Method (Visio)

Returns the path to the specified custom stencil used to populate certain galleries in the Microsoft Visio user interface.


## Syntax

 _expression_ . **GetCustomStencilFile**( **_StencilType_** )

 _expression_ A variable that represents an **[Application](application-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StencilType_|Required| **[VisBuiltInStencilTypes](visbuiltinstenciltypes-enumeration-visio.md)**|The stencil to retrieve. See Remarks for possible values.|

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

