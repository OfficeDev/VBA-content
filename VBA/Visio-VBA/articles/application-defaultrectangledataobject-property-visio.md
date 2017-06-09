---
title: Application.DefaultRectangleDataObject Property (Visio)
keywords: vis_sdr.chm10060050
f1_keywords:
- vis_sdr.chm10060050
ms.prod: visio
api_name:
- Visio.Application.DefaultRectangleDataObject
ms.assetid: 22e7f5ff-516d-4bd0-82bf-2363d1cad973
ms.date: 06/08/2017
---


# Application.DefaultRectangleDataObject Property (Visio)

Returns an  **IDataObject** interface that represents the **Rectangle** tool used in the Microsoft Visio user interface. Read-only.


## Syntax

 _expression_ . **DefaultRectangleDataObject**

 _expression_ An expression that returns a **Application** object.


### Return Value

IDataObject


## Remarks

By using the  **DefaultRectangleDataObject** property to get an **IDataObject** interface, you can create a new rectangle shape linked to data—a result similar to that you would get by dragging a data recordset row onto the page. This property is useful in situations where no master is selected in a docked stencil.


