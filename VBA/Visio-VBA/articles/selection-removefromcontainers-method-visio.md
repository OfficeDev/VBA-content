---
title: Selection.RemoveFromContainers Method (Visio)
keywords: vis_sdr.chm11162220
f1_keywords:
- vis_sdr.chm11162220
ms.prod: visio
api_name:
- Visio.Selection.RemoveFromContainers
ms.assetid: d1ed1360-3caa-3e03-98ef-84f4bd52a035
ms.date: 06/08/2017
---


# Selection.RemoveFromContainers Method (Visio)

Removes the selection of shapes from all lists and containers of which the selection is a member.


## Syntax

 _expression_ . **RemoveFromContainers**

 _expression_ A variable that represents a **[Selection](selection-object-visio.md)** object.


### Return Value

 **Nothing**


## Remarks

When you call the  **RemoveFromContainers** method, Microsoft Visio uses the setting of the **[ContainerProperties.ResizeAsNeeded](containerproperties-resizeasneeded-property-visio.md)** property for each container to determine how the container resizes.


