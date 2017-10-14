---
title: ContainerProperties.HeadingStyle Property (Visio)
keywords: vis_sdr.chm17662620
f1_keywords:
- vis_sdr.chm17662620
ms.prod: visio
api_name:
- Visio.ContainerProperties.HeadingStyle
ms.assetid: aeb0a6c8-fa7d-fe16-a756-84d092d372c1
ms.date: 06/08/2017
---


# ContainerProperties.HeadingStyle Property (Visio)

Determines the appearance and position of the heading of the container. Read/write.


## Syntax

 _expression_ . **HeadingStyle**

 _expression_ An expression that returns a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Return Value

 **Integer**


## Remarks

The value of the  **HeadingStyle** property corresponds to the numerical identifier of the heading style that is selected in the **Heading Styles** gallery in the **Container Styles** group on the **Container Tools Format** tab.

The value of the  **HeadingStyle** should always be greater than or equal to zero (0). A value of zero means that the container does not display a heading.

If no value is assigned to the  **HeadingStyle** property or it is set to a null value, a runtime error ensues. A runtime error also ensues if you assign the property a value less than 0 or greater than the maximum ID number in the **Heading Styles** gallery.


