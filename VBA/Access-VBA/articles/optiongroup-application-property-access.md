---
title: OptionGroup.Application Property (Access)
keywords: vbaac10.chm10809
f1_keywords:
- vbaac10.chm10809
ms.prod: access
api_name:
- Access.OptionGroup.Application
ms.assetid: 21d0325e-4552-699e-4972-1fc5ee157b21
ms.date: 06/08/2017
---


# OptionGroup.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents an **OptionGroup** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[OptionGroup Object](optiongroup-object-access.md)

