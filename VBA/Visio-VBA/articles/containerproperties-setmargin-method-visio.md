---
title: ContainerProperties.SetMargin Method (Visio)
keywords: vis_sdr.chm17662305
f1_keywords:
- vis_sdr.chm17662305
ms.prod: visio
api_name:
- Visio.ContainerProperties.SetMargin
ms.assetid: 008dbfe9-53d9-17a6-c441-b30d5a691716
ms.date: 06/08/2017
---


# ContainerProperties.SetMargin Method (Visio)

Sets the gap between the container and member shapes to the specified size, in the specified units.


## Syntax

 _expression_ . **SetMargin**( **_MarginUnits_** , **_MarginSize_** )

 _expression_ A variable that represents a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MarginUnits_|Required| **[VisUnitCodes](visunitcodes-enumeration-visio.md)**|The units in which to measure the margin.|
| _MarginSize_|Required| **Double**|The size of the margin.|

### Return Value

 **Nothing**


