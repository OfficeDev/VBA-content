---
title: Application.DeferRelationshipRecalc Property (Visio)
keywords: vis_sdr.chm10062425
f1_keywords:
- vis_sdr.chm10062425
ms.prod: visio
api_name:
- Visio.Application.DeferRelationshipRecalc
ms.assetid: b85ce4e4-4425-e508-042f-4119353a60b8
ms.date: 06/08/2017
---


# Application.DeferRelationshipRecalc Property (Visio)

Determines whether Microsoft Visio defers recalculating shape sizes and relationships when a member of the relationship pair is moved or resized. Read/write.


## Syntax

 _expression_ . **DeferRelationshipRecalc**

 _expression_ A variable that represents an **[Application](application-object-visio.md)** object.


### Return Value

 **Boolean**


## Remarks

For example, if you resize a shape that is a member of a container in a structured diagram, Visio will not adjust the size of the container if  **DeferRelationshipRecalc** is **True** . When you set **DeferRelationshipRecalc** to **False** , Visio recalculates the container size and adjusts it accordingly. (In each case, the container's **[ResizeAsNeeded](containerproperties-resizeasneeded-property-visio.md)** property must be set to **visContainerAutoResizeExpandContract** .)

Setting  **DeferRelationshipRecalc** to **False** causes Visio to immediately process all deferred actions.


