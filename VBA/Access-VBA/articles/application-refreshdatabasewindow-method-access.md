---
title: Application.RefreshDatabaseWindow Method (Access)
keywords: vbaac10.chm12563
f1_keywords:
- vbaac10.chm12563
ms.prod: access
api_name:
- Access.Application.RefreshDatabaseWindow
ms.assetid: 63825d35-b24e-ae68-3214-5727dc97eb79
ms.date: 06/08/2017
---


# Application.RefreshDatabaseWindow Method (Access)

The  **RefreshDatabaseWindow** method updates the Database window after a database object has been created, deleted, or renamed.


## Syntax

 _expression_. **RefreshDatabaseWindow**

 _expression_ A variable that represents an **Application** object.


### Return Value

Nothing


## Remarks

You can use the  **RefreshDatabaseWindow** method to immediately reflect changes to objects in Microsoft Access in the Database window. For example, if you add a new form from Visual Basic and save it, you can use the **RefreshDatabaseWindow** method to display the name of the new form on the **Forms** tab of the Database window immediately after it has been saved.


## Example

The following example creates a new form, saves it, and refreshes the Database window:


```vb
Sub CreateFormAndRefresh() 
 Dim frm As Form 
 
 Set frm = CreateForm 
 DoCmd.Save , "NewForm" 
 RefreshDatabaseWindow 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

