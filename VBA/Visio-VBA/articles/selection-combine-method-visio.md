---
title: Selection.Combine Method (Visio)
keywords: vis_sdr.chm11116130
f1_keywords:
- vis_sdr.chm11116130
ms.prod: visio
api_name:
- Visio.Selection.Combine
ms.assetid: a74b25b0-6957-2088-f34f-4000c2be9736
ms.date: 06/08/2017
---


# Selection.Combine Method (Visio)

Creates a new shape by combining selected shapes.


## Syntax

 _expression_ . **Combine**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Remarks

Calling the  **Combine** method is equivalent to clicking **Combine** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab). The produced shape will be the topmost shape in its containing shape and will inherit the text and formatting of the first selected shape. The original shapes are deleted and no shapes are selected when the operation is complete.

The  **Combine** method is similar to the **Join** method but differs in the following ways:




- The  **Combine** method produces a shape with one Geometry section for each original shape. The resulting shape will have holes in regions where the original shapes overlapped.
    
- The  **Join** method differs from **Combine** in that it will coalesce abutting line and curve segments in the original shapes into a single Geometry section in the resulting shape.
    



