---
title: InvisibleApp.DefaultRectangleDataObject Property (Visio)
keywords: vis_sdr.chm17560050
f1_keywords:
- vis_sdr.chm17560050
ms.prod: visio
api_name:
- Visio.InvisibleApp.DefaultRectangleDataObject
ms.assetid: 3ffedd3b-e769-a8a3-e6c0-0d75f7187466
ms.date: 06/08/2017
---


# InvisibleApp.DefaultRectangleDataObject Property (Visio)

Returns an  **IDataObject** interface that represents the **Rectangle** tool used in the Microsoft Visio user interface. Read-only.


## Syntax

 _expression_ . **DefaultRectangleDataObject**

 _expression_ An expression that returns a **InvisibleApp** object.


### Return Value

IDataObject


## Remarks

By using the  **DefaultRectangleDataObject** property to get an **IDataObject** interface, you can create a new rectangle shape linked to data—a result similar to that you would get by dragging a data recordset row onto the page. This property is useful in situations where no master is selected in a docked stencil.


