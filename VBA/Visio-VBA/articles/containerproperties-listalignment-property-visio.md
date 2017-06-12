---
title: ContainerProperties.ListAlignment Property (Visio)
keywords: vis_sdr.chm17662595
f1_keywords:
- vis_sdr.chm17662595
ms.prod: visio
api_name:
- Visio.ContainerProperties.ListAlignment
ms.assetid: f8d62807-9663-b5ac-0154-d37fea1f9816
ms.date: 06/08/2017
---


# ContainerProperties.ListAlignment Property (Visio)

Specifies how to align and arrange a list shape that you want positioned perpendicular to the main list direction. Read/write.


## Syntax

 _expression_ . **ListAlignment**

 _expression_ An expression that returns a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Return Value

 **[VisListAlignment](vislistalignment-enumeration-visio.md)**


## Remarks

Use the  **ListAlignment** property to position shapes along the axis that is perpendicular to the primary list direction. For example, if the primary list direction is horizontal in a given list container, you can use the **ListAlignment** property to align a shape vertically in that container. The value of the **ListAlignment** property can be one of the following **VisListAlignment** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visListAlignLeftOrTop**|0|Left-align or top-align shapes.|
| **visListDirCenterOrMiddle**|1|Center-align or middle-align shapes.|
| **visListDirRightOrBottom**|2|Right-align or bottom-align shapes.|
If the container is not a list, Microsoft Visio returns an  **Invalid Source** error.


