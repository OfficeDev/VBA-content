---
title: UserPermission.UserId Property (Office)
keywords: vbaof11.chm260001
f1_keywords:
- vbaof11.chm260001
ms.prod: office
api_name:
- Office.UserPermission.UserId
ms.assetid: 63c7f01b-3b41-6245-7d3f-5c6440703ccf
ms.date: 06/08/2017
---


# UserPermission.UserId Property (Office)

Gets the e-mail name of the user whose permissions on the active document are determined by the specified  **[UserPermission](userpermission-object-office.md)** object. Read-only.


## Syntax

 _expression_. **UserId**

 _expression_ A variable that represents a **UserPermission** object.


## Remarks

The  **UserPermission** object associates a set of permissions on the active document with a single user and an optional expiration date. The **UserID** property returns the name in e-mail form of the user whose permissions are determined by the specified **UserPermission** object.


## Example

The following example lists the users who have permissions on the active document.


```
 Dim irmPermission As Office.Permission 
 Dim irmUserPerm As Office.UserPermission 
 Dim strUsers As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 For Each irmUserPerm In irmPermission 
 strUsers = strUsers &amp; irmUserPerm.UserId &amp; vbCrLf 
 Next 
 MsgBox strUsers, _ 
 vbInformation + vbOKOnly, "IRM Information" 
 Else 
 MsgBox "Permissions are not enabled for this document.", _ 
 vbInformation + vbOKOnly, "IRM Information" 
 End If 
 Set irmUserPerm = Nothing 
 Set irmPermission = Nothing 

```


## See also


#### Concepts


[UserPermission Object](userpermission-object-office.md)
#### Other resources


[UserPermission Object Members](userpermission-members-office.md)

