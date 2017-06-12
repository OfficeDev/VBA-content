---
title: ComboBox.Application Property (Access)
keywords: vbaac10.chm11356
f1_keywords:
- vbaac10.chm11356
ms.prod: access
api_name:
- Access.ComboBox.Application
ms.assetid: 21c195f2-7a1f-a945-504e-6c1a7fa7f01f
ms.date: 06/08/2017
---


# ComboBox.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

