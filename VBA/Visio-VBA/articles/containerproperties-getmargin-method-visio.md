---
title: ContainerProperties.GetMargin Method (Visio)
keywords: vis_sdr.chm17662300
f1_keywords:
- vis_sdr.chm17662300
ms.prod: visio
api_name:
- Visio.ContainerProperties.GetMargin
ms.assetid: c0e224a1-f7a6-e16c-a99c-766a5a4ac207
ms.date: 06/08/2017
---


# ContainerProperties.GetMargin Method (Visio)

Returns the minimal distance, in the specified units, between the edges of the container or list and those of its member shapes.


## Syntax

 _expression_ . **GetMargin**( **_MarginUnits_** )

 _expression_ A variable that represents a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MarginUnits_|Required| **[VisUnitCodes](visunitcodes-enumeration-visio.md)**|The units in which to measure the margin.|

### Return Value

 **Double**


