---
title: Master.AlignName Property (Visio)
keywords: vis_sdr.chm10713075
f1_keywords:
- vis_sdr.chm10713075
ms.prod: visio
api_name:
- Visio.Master.AlignName
ms.assetid: 5df055eb-ddb1-2d2a-1d94-93781960b3a9
ms.date: 06/08/2017
---


# Master.AlignName Property (Visio)

Gets or sets the position of a master name in a stencil window. Read/write.


## Syntax

 _expression_ . **AlignName**

 _expression_ A variable that represents a **Master** object.


### Return Value

Integer


## Remarks

Only user-created stencils are editable. By default, Visio stencils are not editable. 

The following constants declared by the Visio type library show the possible alignment values.



|**Constant **|**Value **|
|:-----|:-----|
| **visLeft**|1|
| **visCenter**|2|
| **visRight**|3|

