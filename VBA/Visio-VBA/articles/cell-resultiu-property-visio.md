---
title: Cell.ResultIU Property (Visio)
keywords: vis_sdr.chm10114220
f1_keywords:
- vis_sdr.chm10114220
ms.prod: visio
api_name:
- Visio.Cell.ResultIU
ms.assetid: 4d752d78-e112-bb45-08c7-5411d7d79beb
ms.date: 06/08/2017
---


# Cell.ResultIU Property (Visio)

Gets or sets a cell's value in internal units. Read/write.


## Syntax

 _expression_ . **ResultIU**

 _expression_ A variable that represents a **Cell** object.


### Return Value

Double


## Remarks

Use the  **ResultIU** property to set the value of an unguarded cell. If a cell's formula is protected with a GUARD function, the formula is not changed and an error is generated. To set the value of a guarded cell in internal units, use the **ResultIUForce** property.

The units default to the Microsoft Visio internal units, which are inches for distance and radians for angles.


