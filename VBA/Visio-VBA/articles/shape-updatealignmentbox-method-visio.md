---
title: Shape.UpdateAlignmentBox Method (Visio)
keywords: vis_sdr.chm11216635
f1_keywords:
- vis_sdr.chm11216635
ms.prod: visio
api_name:
- Visio.Shape.UpdateAlignmentBox
ms.assetid: 7076ee5f-f536-77ec-a1f7-518195e3e897
ms.date: 06/08/2017
---


# Shape.UpdateAlignmentBox Method (Visio)

Updates the alignment box for a shape.


## Syntax

 _expression_ . **UpdateAlignmentBox**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Nothing


## Remarks

The  **UpdateAlignmentBox** method alters the width and height of a shape, often a group. For example, after you move a shape in a group, the shape may be outside the group's alignment box. The **UpdateAlignmentBox** method updates the alignment box so that it encloses all the shapes in the group.


 **Note**  Many shapes are designed so that their alignment boxes don't coincide with their geometric extents. Using the  **UpdateAlignmentBox** method on such shapes defeats the intentions of the shape designer.


