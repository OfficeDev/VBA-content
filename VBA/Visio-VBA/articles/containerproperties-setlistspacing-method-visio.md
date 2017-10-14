---
title: ContainerProperties.SetListSpacing Method (Visio)
keywords: vis_sdr.chm17662315
f1_keywords:
- vis_sdr.chm17662315
ms.prod: visio
api_name:
- Visio.ContainerProperties.SetListSpacing
ms.assetid: 2aa7d9c3-5945-5b2c-ab0c-3663e6d49288
ms.date: 06/08/2017
---


# ContainerProperties.SetListSpacing Method (Visio)

Sets the gap between adjacent member shapes in the list.


## Syntax

 _expression_ . **SetListSpacing**( **_SpacingUnits_** , **_SpacingSize_** )

 _expression_ A variable that represents a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SpacingUnits_|Required| **[VisUnitCodes](visunitcodes-enumeration-visio.md)**|The units in which to measure the gap.|
| _SpacingSize_|Required| **Double**|The size of the gap.|

### Return Value

 **Nothing**


## Remarks

If the container is not a list, Microsoft Visio returns an Invalid Source error.


