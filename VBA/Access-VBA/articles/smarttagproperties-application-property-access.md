---
title: SmartTagProperties.Application Property (Access)
keywords: vbaac10.chm13310
f1_keywords:
- vbaac10.chm13310
ms.prod: access
api_name:
- Access.SmartTagProperties.Application
ms.assetid: 4a282407-1dc4-1a21-41b3-f7601eb59dfc
ms.date: 06/08/2017
---


# SmartTagProperties.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **SmartTagProperties** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[SmartTagProperties Collection](smarttagproperties-object-access.md)

