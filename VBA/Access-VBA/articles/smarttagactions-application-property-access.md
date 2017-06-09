---
title: SmartTagActions.Application Property (Access)
keywords: vbaac10.chm13297
f1_keywords:
- vbaac10.chm13297
ms.prod: access
api_name:
- Access.SmartTagActions.Application
ms.assetid: 51c4f3b3-e1a9-2f69-146a-2d9d2cac7e5c
ms.date: 06/08/2017
---


# SmartTagActions.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **SmartTagActions** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[SmartTagActions Collection](smarttagactions-object-access.md)

