---
title: BoundObjectFrame.Application Property (Access)
keywords: vbaac10.chm10896
f1_keywords:
- vbaac10.chm10896
ms.prod: access
api_name:
- Access.BoundObjectFrame.Application
ms.assetid: 05b5b479-fe8b-6d03-b8de-59afa7a587b9
ms.date: 06/08/2017
---


# BoundObjectFrame.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

