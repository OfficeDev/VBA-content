---
title: Shape.LengthIU Property (Visio)
keywords: vis_sdr.chm11213835
f1_keywords:
- vis_sdr.chm11213835
ms.prod: visio
api_name:
- Visio.Shape.LengthIU
ms.assetid: 11d57f17-5285-6b45-1da1-dc58db087395
ms.date: 06/08/2017
---


# Shape.LengthIU Property (Visio)

Returns the length (perimeter) of the shape in internal units. Read-only.


## Syntax

 _expression_ . **LengthIU**( **_fIncludeSubShapes_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fIncludeSubShapes_|Optional| **Boolean**| **False** to exclude the length of subshapes. Length of subshapes is included by default.|

### Return Value

Double


## Remarks

Data graphic callout shapes (and their sub-shapes) that are applied to the parent shape are excluded from length calculations. If the parent shape is itself a data graphic callout shape, its geometry (and that of its sub-shapes) is  _not_ excluded from length calculations.

In versions before Microsoft Office Visio 2003, this property took no arguments.


