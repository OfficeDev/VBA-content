---
title: Shape.RemoveFromContainers Method (Visio)
keywords: vis_sdr.chm11262220
f1_keywords:
- vis_sdr.chm11262220
ms.prod: visio
api_name:
- Visio.Shape.RemoveFromContainers
ms.assetid: b9dbf604-01f0-675a-a0e1-7b30841ec5c5
ms.date: 06/08/2017
---


# Shape.RemoveFromContainers Method (Visio)

Removes the shape from all lists and containers of which it is a member.


## Syntax

 _expression_ . **RemoveFromContainers**

 _expression_ A variable that represents a **[Shape](shape-object-visio.md)** object.


### Return Value

 **Nothing**


## Remarks

When you call the  **RemoveFromContainers** method, Microsoft Visio uses the **[ContainerProperties.ResizeAsNeeded](containerproperties-resizeasneeded-property-visio.md)** property setting for each container to determine how to resize the container.


