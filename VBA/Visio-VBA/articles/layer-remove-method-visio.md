---
title: Layer.Remove Method (Visio)
keywords: vis_sdr.chm11816470
f1_keywords:
- vis_sdr.chm11816470
ms.prod: visio
api_name:
- Visio.Layer.Remove
ms.assetid: d46c814b-1937-de81-de1b-e670667920c2
ms.date: 06/08/2017
---


# Layer.Remove Method (Visio)

Removes a shape from a layer.


## Syntax

 _expression_ . **Remove**( **_SheetObject_** , **_fPresMems_** )

 _expression_ A variable that represents a **Layer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SheetObject_|Required| **[IVSHAPE]**|An expression that returns the  **Shape** object to remove.|
| _fPresMems_|Required| **Integer**|Flag that indicates whether to remove members of a group.|

### Return Value

Nothing


## Remarks

If the shape is a group and  _fPresMems_ is non-zero, member shapes of the group are unaffected. If _fPresMems_ is zero (0), the group's member shapes are also removed from the layer.

Removing a shape from a layer does not delete the shape.


