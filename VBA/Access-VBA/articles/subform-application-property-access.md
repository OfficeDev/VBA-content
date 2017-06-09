---
title: SubForm.Application Property (Access)
keywords: vbaac10.chm11914
f1_keywords:
- vbaac10.chm11914
ms.prod: access
api_name:
- Access.SubForm.Application
ms.assetid: 2aafea49-e27c-b3d4-2710-b1ef1c84b195
ms.date: 06/08/2017
---


# SubForm.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **SubForm** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[SubForm Object](subform-object-access.md)

