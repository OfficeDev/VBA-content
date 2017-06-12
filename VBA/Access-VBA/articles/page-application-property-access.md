---
title: Page.Application Property (Access)
keywords: vbaac10.chm12138
f1_keywords:
- vbaac10.chm12138
ms.prod: access
api_name:
- Access.Page.Application
ms.assetid: eaae5c67-91af-ac9d-685c-81f82ec2b023
ms.date: 06/08/2017
---


# Page.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **Page** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[Page Object](page-object-access.md)

