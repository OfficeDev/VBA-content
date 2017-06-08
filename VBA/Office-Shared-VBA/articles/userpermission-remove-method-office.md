---
title: UserPermission.Remove Method (Office)
keywords: vbaof11.chm260005
f1_keywords:
- vbaof11.chm260005
ms.prod: office
api_name:
- Office.UserPermission.Remove
ms.assetid: d4c8778f-dc1b-7d5b-6a7a-65b91909bfe3
ms.date: 06/08/2017
---


# UserPermission.Remove Method (Office)

Removes the specified  **UserPermission** object from the **[Permission](permission-object-office.md)** collection of the active document.


## Syntax

 _expression_. **Remove**

 _expression_ Required. A variable that represents a **[UserPermission](userpermission-object-office.md)** object.


## Remarks

The  **UserPermission** object associates a set of permissions on the active document with a single user and an optional expiration date. The **Remove** method removes the user and the set of user permissions determined by the specified **UserPermission** object.


## Example

The following example removes the second user's permissions on the active document from the document's Permission collection.


```
 Dim irmPermission As Office.Permission 
 Dim irmUserPerm As Office.UserPermission 
 Set irmPermission = ActiveWorkbook.Permission 
 Set irmUserPerm = irmPermission.Item(2) 
 irmUserPerm.Remove 
 MsgBox "Permission removed.", _ 
 vbInformation + vbOKOnly, "IRM Information" 
 Set irmUserPerm = Nothing 
 Set irmPermission = Nothing 

```


## See also


#### Concepts


[UserPermission Object](userpermission-object-office.md)
#### Other resources


[UserPermission Object Members](userpermission-members-office.md)

