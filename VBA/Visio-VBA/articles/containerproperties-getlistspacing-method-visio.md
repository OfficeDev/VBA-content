---
title: ContainerProperties.GetListSpacing Method (Visio)
keywords: vis_sdr.chm17662310
f1_keywords:
- vis_sdr.chm17662310
ms.prod: visio
api_name:
- Visio.ContainerProperties.GetListSpacing
ms.assetid: cc20b7dc-1498-998d-23fa-a69bbba35294
ms.date: 06/08/2017
---


# ContainerProperties.GetListSpacing Method (Visio)

Returns the gap between adjacent member shapes in the list.


## Syntax

 _expression_ . **GetListSpacing**( **_SpacingUnits_** )

 _expression_ A variable that represents a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SpacingUnits_|Required| **[VisUnitCodes](visunitcodes-enumeration-visio.md)**|The units in which to measure the gap.|

### Return Value

 **Double**


## Remarks

If the container is not a list, Microsoft Visio returns an Invalid Source error.


