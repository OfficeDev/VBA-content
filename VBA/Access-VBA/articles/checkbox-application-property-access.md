---
title: CheckBox.Application Property (Access)
keywords: vbaac10.chm10682
f1_keywords:
- vbaac10.chm10682
ms.prod: access
api_name:
- Access.CheckBox.Application
ms.assetid: b47ba26d-bd87-ec43-381a-c93b654c08e2
ms.date: 06/08/2017
---


# CheckBox.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **CheckBox** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[CheckBox Object](checkbox-object-access.md)

