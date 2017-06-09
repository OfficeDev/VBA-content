---
title: Application.CommandBars Property (Access)
keywords: vbaac10.chm12559
f1_keywords:
- vbaac10.chm12559
ms.prod: access
api_name:
- Access.Application.CommandBars
ms.assetid: a7dc2e41-7271-1f2d-b0f9-7fa884311709
ms.date: 06/08/2017
---


# Application.CommandBars Property (Access)

You can use the  **CommandBars** property to return a reference to the **CommandBars** collection object. Read-only **CommandBars** object.


## Syntax

 _expression_. **CommandBars**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **CommandBars** collection object is the collection of all built-in and custom command bars in an application. You can refer to individual members of the collection by using the member object's index or a string expression that is the name of the member object. The first member object in the collection has an index value of 1 and the total number of member objects in the collection is the value of the **CommandBars** collection's **Count** property.

Once you establish a reference to the  **CommandBars** collection object, you can access all the properties and methods of the object. You can set a reference to the **CommandBars** collection object by clicking **References** on the **Tools** menu while in module Design view. Set a reference to the Microsoft Office 12.0 Object Library in the **References** dialog box by selecting the appropriate check box.


## See also


#### Concepts


[Application Object](application-object-access.md)

