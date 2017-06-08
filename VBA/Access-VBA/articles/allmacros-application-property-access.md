---
title: AllMacros.Application Property (Access)
keywords: vbaac10.chm12678
f1_keywords:
- vbaac10.chm12678
ms.prod: access
api_name:
- Access.AllMacros.Application
ms.assetid: 8297f6a4-fbfb-9d73-1914-6bc351f09619
ms.date: 06/08/2017
---


# AllMacros.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents an **AllMacros** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[AllMacros Collection](allmacros-object-access.md)

