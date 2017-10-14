---
title: Selection.Intersect Method (Visio)
keywords: vis_sdr.chm11116375
f1_keywords:
- vis_sdr.chm11116375
ms.prod: visio
api_name:
- Visio.Selection.Intersect
ms.assetid: 5dc63a77-62de-3892-6ed2-bcb5cb0a29f1
ms.date: 06/08/2017
---


# Selection.Intersect Method (Visio)

Creates one closed shape from the area in which selected shapes overlap or intersect.


## Syntax

 _expression_ . **Intersect**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Remarks

Calling the  **Intersect** method is equivalent to clicking **Intersect** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab). The produced shape will be the topmost shape in its containing shape and will inherit the text and formatting of the first selected shape.

The original shapes are deleted and no shapes are selected when the operation is complete.


