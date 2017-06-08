---
title: ContainerProperties.ContainerStyle Property (Visio)
keywords: vis_sdr.chm17651150
f1_keywords:
- vis_sdr.chm17651150
ms.prod: visio
api_name:
- Visio.ContainerProperties.ContainerStyle
ms.assetid: cc7b6757-0287-e25a-9406-554aa70ef181
ms.date: 06/08/2017
---


# ContainerProperties.ContainerStyle Property (Visio)

Determines the visual appearance of the container. Read/write.


## Syntax

 _expression_ . **ContainerStyle**

 _expression_ An expression that returns a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Return Value

 **Integer**


## Remarks

The value of the  **ContainerStyle** property corresponds to the numerical identifier (ID) of the container style that is selected in the **Container Styles** gallery on the **Container Tools Format** tab.

The value of the  **ContainerStyle** should always be greater than zero.

If no value is assigned to the  **ContainerStyle** property or it is set to a null value, a runtime error ensues. A runtime error also ensues if you assign the property a value that is less than 1 or greater than the maximum ID number in the **Container Styles** gallery.


