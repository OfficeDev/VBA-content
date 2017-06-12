---
title: Master.MatchByName Property (Visio)
keywords: vis_sdr.chm10713890
f1_keywords:
- vis_sdr.chm10713890
ms.prod: visio
api_name:
- Visio.Master.MatchByName
ms.assetid: 4edb0e5f-7e87-c66d-b842-318cd0eba5d5
ms.date: 06/08/2017
---


# Master.MatchByName Property (Visio)

Determines how the application decides if a document master is already present when an instance of a master is dropped on the drawing page. It allows changes made to a document master to apply to new instances of the master, even if the instances are dragged from a stand-alone stencil file. Read/write.


## Syntax

 _expression_ . **MatchByName**

 _expression_ A variable that represents a **Master** object.


### Return Value

Integer


## Remarks

Setting the  **MatchByName** property is equivalent to selecting or clearing the **Match master by name on drop** check box in the **Master Properties** dialog box (right-click the master, point to **Edit Master**, and then click  **Master Properties** on the shortcut menu).

Suppose you create an instance of a shape master from a stand-alone stencil (producing a local copy of the shape master in the document stencil) and then make modifications to the new shape master instance (such as changing its fill color): 

- If the  **MatchByName** property of the document master is **False** , dragging the original shape master from the stand-alone stencil into the drawing creates an instance that has the stand-alone master's attributes and produces a new document shape master in the document stencil.

- If the  **MatchByName** property of the document master is **True** , dragging the original master from the stand-alone stencil into the drawing creates a shape instance that has the document master's attributes and doesn't produce a shape _master_ in the document stencil.



