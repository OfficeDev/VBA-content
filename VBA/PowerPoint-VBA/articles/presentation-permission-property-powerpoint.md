---
title: Presentation.Permission Property (PowerPoint)
keywords: vbapp10.chm583082
f1_keywords:
- vbapp10.chm583082
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Permission
ms.assetid: 3f7633a8-bdab-b08d-0cf8-8df52c35865a
ms.date: 06/08/2017
---


# Presentation.Permission Property (PowerPoint)





## Syntax

 _expression_. **Permission**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Permission


## Remarks

Use the  **Permission** object to restrict permissions to the active document and to return or set specific permissions settings.

Use the  **Enabled** property to determine whether permissions are restricted on the active document. Use the **Count** property to return the number of users with permissions, and the **RemoveAll** method to reset all existing permissions.

The  **DocumentAuthor**, **EnableTrustedBrowser**, **RequestPermissionURL**, and **StoreLicenses** properties provide additional information about permission settings.

The  **Permission** object gives access to a collection of **UserPermission** objects. Use the **UserPermission** object to associate specific sets of rights with individual users. While some permissions granted in the user interface (such as **msoPermissionPrint** ) apply to all users, you can use the **UserPermission** object to assign them on a per-user basis with per-user expiration dates.

Information Rights Management supports the use of administrative permission policies, which list users and groups and their document permissions. Use the  **ApplyPolicy** method to apply a permission policy, and the **PermissionFromPolicy**, **PolicyName**, and **PolicyDescription** properties to return policy information.

The  **Permission** object model is available whether permissions are restricted on the active document or not. The **Permission** property of the **Presentation** object does not return **Nothing** when the active document does not have restricted permissions. Use the **Enabled** property to determine whether a document has restricted permissions.


## Example

The following example creates a new presentation and assigns the user with e-mail address "someone@example.com" read permission on the new presentation. The example will display the permissions of the owner and the new user.


```vb
Sub AddUserPermissions()

 Dim myPres As PowerPoint.Presentation

 Dim myPer As Office.Permission

 Dim NewOwnerPer As Office.UserPermission

 Set myPres = Application.Presentations.Add(msoTrue)

 Set myPer = myPres.Permission

 myPer.Enabled = True

 Set NewOwnerPer = myPer.Add("someone@example.com", msoPermissionRead )

 MsgBox myPer(1).UserId + " " + Str(myPer(1).Permission)

 MsgBox myPer(2).UserId + " " + Str(myPer(2).Permission)

End Sub
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

