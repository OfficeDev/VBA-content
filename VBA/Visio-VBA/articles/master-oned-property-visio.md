---
title: Master.OneD Property (Visio)
keywords: vis_sdr.chm10713975
f1_keywords:
- vis_sdr.chm10713975
ms.prod: visio
api_name:
- Visio.Master.OneD
ms.assetid: 917f8cfc-a2fc-7572-936a-69956d139131
ms.date: 06/08/2017
---


# Master.OneD Property (Visio)

Determines whether an object behaves as a one-dimensional (1-D) object. Read-only.


## Syntax

 _expression_ . **OneD**

 _expression_ A variable that represents a **Master** object.


### Return Value

Integer


## Remarks

Setting the  **OneD** property is equivalent to changing a shape's interaction style in the **Behavior** dialog box (click **Behavior** in the **Shape Design** group of the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab). Setting the **OneD** property for a 1-D shape to **False** deletes its 1-D Endpoints section, even if the cells in that section were protected with the GUARD function.

You can get, but not set, the  **OneD** property of a **Master** object.

The  **OneD** property of a **Shape** object that is a guide is always 0. If you try to change the value of the **OneD** property of a guide shape, no error is raised, but the value remains 0.

The  **OneD** property of an object from another application is always **False** .


