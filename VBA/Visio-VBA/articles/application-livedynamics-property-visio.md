---
title: Application.LiveDynamics Property (Visio)
keywords: vis_sdr.chm10013855
f1_keywords:
- vis_sdr.chm10013855
ms.prod: visio
api_name:
- Visio.Application.LiveDynamics
ms.assetid: fc5a887b-318a-fd25-c2b5-52d6cc1c026e
ms.date: 06/08/2017
---


# Application.LiveDynamics Property (Visio)

Controls whether Microsoft Visio recalculates shape properties during drag operations on every mouse move or only after the mouse button is released. Read/write.


## Syntax

 _expression_ . **LiveDynamics**

 _expression_ A variable that represents an **Application** object.


### Return Value

Boolean


## Remarks

The  **LiveDynamics** property tracks actions, such as resizing and rotating shapes, and is effective when shapes are glued or related to each other. When the value of the **LiveDynamics** property is **True** , more events such as **CellChanged** occur. Solutions that respond to such events may operate more quickly if the **LiveDynamics** property is set to **False** .


