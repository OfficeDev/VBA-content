---
title: Selection.Union Method (Visio)
keywords: vis_sdr.chm11116630
f1_keywords:
- vis_sdr.chm11116630
ms.prod: visio
api_name:
- Visio.Selection.Union
ms.assetid: 1ab7ce2a-98af-c455-7558-6f4f9226eeb9
ms.date: 06/08/2017
---


# Selection.Union Method (Visio)

Creates a new shape from the perimeter of selected shapes.


## Syntax

 _expression_ . **Union**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Remarks

Calling the  **Union** method is equivalent to clicking **Union** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab). The produced shape will be the topmost shape in its containing shape and will inherit the text and formatting of the first selected shape.

When the operation is complete, the original shapes are deleted and no shapes are selected.


