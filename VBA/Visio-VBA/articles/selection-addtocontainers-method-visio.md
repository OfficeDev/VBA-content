---
title: Selection.AddToContainers Method (Visio)
keywords: vis_sdr.chm11162215
f1_keywords:
- vis_sdr.chm11162215
ms.prod: visio
api_name:
- Visio.Selection.AddToContainers
ms.assetid: 7f3e739f-a573-049c-9f54-9e93a401191f
ms.date: 06/08/2017
---


# Selection.AddToContainers Method (Visio)

Adds the selection of shapes to all underlying containers that allow it as a member.


## Syntax

 _expression_ . **AddToContainers**

 _expression_ A variable that represents a **[Selection](selection-object-visio.md)** object.


### Return Value

 **Nothing**


## Remarks

When you call the  **AddToContainers** method, Microsoft Visio uses the setting of the **[ContainerProperties.ResizeAsNeeded](containerproperties-resizeasneeded-property-visio.md)** property for each container to determine how the container resizes.

Each shape in the selection is added to its underlying containers according to the position of the shape. As a result, different shapes may end up being contained by different containers. If the underlying container is a list, the shape is added as normal container member, not list member.

The  **AddToContainers** method works only if the selection sits at least partially on top of a container that does not already contain it.


