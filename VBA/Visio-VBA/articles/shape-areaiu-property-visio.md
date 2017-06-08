---
title: Shape.AreaIU Property (Visio)
keywords: vis_sdr.chm11213095
f1_keywords:
- vis_sdr.chm11213095
ms.prod: visio
api_name:
- Visio.Shape.AreaIU
ms.assetid: a9982cd2-9a91-f5e5-7297-360b6d9a1f29
ms.date: 06/08/2017
---


# Shape.AreaIU Property (Visio)

Returns the area of a  **Shape** object in internal units (square inches). Read-only.


## Syntax

 _expression_ . **AreaIU**( **_fIncludeSubShapes_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fIncludeSubShapes_|Optional| **Boolean**| **False** to exclude the area of subshapes. Area of subshapes is included by default.|

### Return Value

Double


## Remarks

Data graphic callout shapes (and their sub-shapes) that are applied to the parent shape are excluded from area calculations. If the parent shape is itself a data graphic callout shape, its geometry (and that of its sub-shapes) is  _not_ excluded from area calculations.

In versions before Microsoft Office Visio 2003, this property took no arguments.


