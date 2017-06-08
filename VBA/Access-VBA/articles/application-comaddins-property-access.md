---
title: Application.COMAddIns Property (Access)
keywords: vbaac10.chm12585
f1_keywords:
- vbaac10.chm12585
ms.prod: access
api_name:
- Access.Application.COMAddIns
ms.assetid: b94474b4-3690-54ab-1a4b-b30744354db5
ms.date: 06/08/2017
---


# Application.COMAddIns Property (Access)

You can use the  **COMAddIns** property to return a reference to the current **COMAddIns** collection object and its related properties. Read-only **COMAddIns** object.


## Syntax

 _expression_. **COMAddIns**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **COMAddIns** collection object is the collection of all currently registered COM add-ins of an application. You can refer to individual members of the collection by using the member object's index or a string expression that is the name of the member object. The first member object in the collection has an index value of 1 and the total number of member objects in the collection is the value of the **COMAddIns** collection's **Count** property.

Once you establish a reference to the  **COMAddIns** collection object, you can access all the properties and methods of the object. You can set a reference to the **COMAddIns** collection object by clicking **References** on the **Tools** menu while in module Design view. Then set a reference to the Microsoft Office 12.0 Object Library in the **References** dialog box by selecting the appropriate check box. Microsoft Access can set this reference for you if you use a Microsoft Office 12.0 Object Library constant to set a **COMAddIns** collection object's property or as an argument to a **COMAddIns** collection object's method.


## See also


#### Concepts


[Application Object](application-object-access.md)

