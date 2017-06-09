---
title: Controls.Application Property (Access)
keywords: vbaac10.chm10177
f1_keywords:
- vbaac10.chm10177
ms.prod: access
api_name:
- Access.Controls.Application
ms.assetid: c8650732-ffee-830b-9d9d-571a09af3a4c
ms.date: 06/08/2017
---


# Controls.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **Controls** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[Controls Collection](controls-object-access.md)

