---
title: Shape.Parent Property (Visio)
keywords: vis_sdr.chm11250755
f1_keywords:
- vis_sdr.chm11250755
ms.prod: visio
api_name:
- Visio.Shape.Parent
ms.assetid: aada0bc1-75e3-8357-3ef9-597a10250860
ms.date: 06/08/2017
---


# Shape.Parent Property (Visio)

Determines the parent of a  **Shape** object. Read/write.


## Syntax

 _expression_ . **Parent**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Object


## Remarks

In general, an object's parent is the object that contains it. If a  **Shape** object is a member of a group, the parent is that group. Otherwise, its parent is a **Page** or a **Master** object.

When assigning a new parent shape, you must assign a  **Shape** object. If you want to assign a page or master to be the parent of a shape, you must assign the **Shape** object returned by the **Page** or **Master** object's **PageSheet** property.

A shape and its parent shape must be in the same containing page or containing master. If the new parent is not a  **Shape** object, or if the **ContainingPage** or **ContainingMaster** property of the parent shape is different from that of the shape, Microsoft Visio raises an exception.


