---
title: Selection.Subtract Method (Visio)
keywords: vis_sdr.chm11116595
f1_keywords:
- vis_sdr.chm11116595
ms.prod: visio
api_name:
- Visio.Selection.Subtract
ms.assetid: 606798b6-3482-0c45-d583-4762ee07da45
ms.date: 06/08/2017
---


# Selection.Subtract Method (Visio)

Subtracts the areas that overlap the selected shape.


## Syntax

 _expression_ . **Subtract**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Remarks

Calling the  **Subtract** method is equivalent to clicking **Subtract** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab). The first selected shape is the one that will have the other selected shapes subtracted from it. The other shapes will be deleted and no shapes are selected when the operation is complete.


