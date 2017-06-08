---
title: OptionButton.Application Property (Access)
keywords: vbaac10.chm10556
f1_keywords:
- vbaac10.chm10556
ms.prod: access
api_name:
- Access.OptionButton.Application
ms.assetid: 8c2e2c14-b66b-435c-4631-d49b8a376671
ms.date: 06/08/2017
---


# OptionButton.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents an **OptionButton** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[OptionButton Object](optionbutton-object-access.md)

