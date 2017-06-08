---
title: Application.UserControl Property (Access)
keywords: vbaac10.chm12514
f1_keywords:
- vbaac10.chm12514
ms.prod: access
api_name:
- Access.Application.UserControl
ms.assetid: e82213ac-bd7b-2669-3001-330f40cfdaaa
ms.date: 06/08/2017
---


# Application.UserControl Property (Access)

You can use the  **UserControl** property to determine whether the current Microsoft Access application was started by the user or by another application with Automation, formerly called OLE Automation. Read/write **Boolean**.


## Syntax

 _expression_. **UserControl**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **UserControl** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|The current application was started by the user.|
|**False**|The current application was started by another application with Automation.|
This property is read-only in all views when user starts the Access application. If Microsoft Access is started by OLE Automation, the  **UserControl** property can be set in Visual Basic.

When an application is launched by the user, the  **Visible** and **UserControl** properties of the **[Application](application-object-access.md)** object are both set to **True**. When the **UserControl** property is set to **True**, it isn't possible to set the **Visible** property of the object to **False**.

When an  **Application** object is created by using Automation, the **Visible** and **UserControl** properties of the object are both set to **False**.


## Example

The following example displays a message indicating whether Access was started by the user.


```vb
MsgBox "The user started Access:  " &; Application.UserControl
```


## See also


#### Concepts


[Application Object](application-object-access.md)

