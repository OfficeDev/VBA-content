---
title: RelatedShapePairEvent Object (Visio)
keywords: vis_sdr.chm61045
f1_keywords:
- vis_sdr.chm61045
ms.prod: visio
api_name:
- Visio.RelatedShapePairEvent
ms.assetid: 8a59ae03-ed45-21e3-73ad-8fdbe4c53299
ms.date: 06/08/2017
---


# RelatedShapePairEvent Object (Visio)

Holds information about the shapes that are involved in a container relationship or a callout relationship.


## Remarks

A related shape pair consists of two shapes — typically a container and a member, or a callout and a target shape.

When you add or remove a member shape from a container, Microsoft Visio fires a  **ContainerRelationshipAdded** or **ContainerRelationshipDeleted** event respectively, and in each case, returns a **RelatedShapePairEvent** object that encapsulates information about the event.

When you add or remove a callout relationship from a document, Microsoft Visio fires a  **CalloutRelationshipAdded** or **CalloutRelationshipDeleted** event respectively, and in each case, returns a **RelatedShapePairEvent** object that encapsulates information about the event.


