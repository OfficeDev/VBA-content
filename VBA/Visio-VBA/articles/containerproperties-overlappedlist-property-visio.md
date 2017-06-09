---
title: ContainerProperties.OverlappedList Property (Visio)
keywords: vis_sdr.chm17662615
f1_keywords:
- vis_sdr.chm17662615
ms.prod: visio
api_name:
- Visio.ContainerProperties.OverlappedList
ms.assetid: e0fb8674-f17d-e48f-b7c4-db11d435dbf4
ms.date: 06/08/2017
---


# ContainerProperties.OverlappedList Property (Visio)

Creates or removes an overlapped list relationship with another list shape, or returns the target list shape that participates in an overlapped list relationship with the source list shape. Read/write.


## Syntax

 _expression_ . **OverlappedList**

 _expression_ An expression that returns a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Return Value

 **[Shape](shape-object-visio.md)**


## Remarks

To create an overlapped list relationship, set  **OverlappedList** equal to the target list shape.

To remove an existing overlapped list relationship between the source list shape and the target list shape, set  **OverlappedList** equal to **Nothing** .

 **OverlappedList** returns **Nothing** if there is no existing overlapped list relationship between the source shape and any other shape.

 **OverlappedList** returns an Invalid Source error if the source shape is not a list. It returns an Invalid Target error if the target shape is not a list.


