---
title: Master.Connects Property (Visio)
keywords: vis_sdr.chm10713290
f1_keywords:
- vis_sdr.chm10713290
ms.prod: visio
api_name:
- Visio.Master.Connects
ms.assetid: 72c01ae0-9134-d384-b860-dbb333a498fe
ms.date: 06/08/2017
---


# Master.Connects Property (Visio)

Returns a  **Connects** collection for a shape, page, or master. Read-only.


## Syntax

 _expression_ . **Connects**

 _expression_ A variable that represents a **Master** object.


### Return Value

Connects


## Remarks

The  **Connects** collection of a shape contains every **Connect** object for which the shape is returned by the **FromSheet** property. This tells you all the shapes to which the shape is connected.

To obtain a  **Connects** collection that contains every **Connect** object for which the shape is the **ToSheet** property, use the shape's **FromConnects** property. This tells you all the shapes that are connected to this shape.

The  **Connects** collection of a page contains a **Connect** object for every connection on the page.

The  **Connects** collection of a master contains a **Connect** object for every connection in the master.


