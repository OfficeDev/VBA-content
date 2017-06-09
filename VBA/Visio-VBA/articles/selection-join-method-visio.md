---
title: Selection.Join Method (Visio)
keywords: vis_sdr.chm11116380
f1_keywords:
- vis_sdr.chm11116380
ms.prod: visio
api_name:
- Visio.Selection.Join
ms.assetid: e176abcc-edd1-0e40-afc8-e05ed8dec998
ms.date: 06/08/2017
---


# Selection.Join Method (Visio)

Creates a new shape by joining selected shapes.


## Syntax

 _expression_ . **Join**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Remarks

Calling the  **Join** method is equivalent to clicking **Join** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab). The new shape inherits the text and formatting of the first selected shape and is the topmost shape in its container—the _n_th shape in the  **Shapes** collection of its containing shape, where _n_ = Count.

The original shapes are deleted and no shapes are selected when the operation is complete.

The  **Join** method and the **Combine** method are similar but differ in the following ways:




-  **Join** coalesces abutting line and curve segments in the original shapes into a single Geometry section in the resulting shape.
    
-  **Combine** produces a shape that has one Geometry section for each original shape. The resulting shape has holes in regions where the original shapes overlapped.
    


You might want to join shapes after importing a non-Visio drawing in which apparent polylines are represented by many independent shapes, each possessing a single line or curve segment. By joining the shapes that constitute a polyline in such a drawing, you can replace many single-segment shapes with one multiple-segment shape.


