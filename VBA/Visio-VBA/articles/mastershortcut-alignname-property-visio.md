---
title: MasterShortcut.AlignName Property (Visio)
keywords: vis_sdr.chm16013075
f1_keywords:
- vis_sdr.chm16013075
ms.prod: visio
api_name:
- Visio.MasterShortcut.AlignName
ms.assetid: 022fdf3a-17f4-740f-191e-a06684ee3112
ms.date: 06/08/2017
---


# MasterShortcut.AlignName Property (Visio)

Gets or sets the position of a master name in a stencil window. Read/write.


## Syntax

 _expression_ . **AlignName**

 _expression_ A variable that represents a **MasterShortcut** object.


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

