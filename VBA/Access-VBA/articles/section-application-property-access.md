---
title: Section.Application Property (Access)
keywords: vbaac10.chm12186
f1_keywords:
- vbaac10.chm12186
ms.prod: access
api_name:
- Access.Section.Application
ms.assetid: 2f3d0784-34a7-b3d2-af29-5ab97f4e4467
ms.date: 06/08/2017
---


# Section.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **Section** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[Section Object](section-object-access.md)

