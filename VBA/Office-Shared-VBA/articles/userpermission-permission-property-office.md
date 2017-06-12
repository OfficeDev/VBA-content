---
title: UserPermission.Permission Property (Office)
keywords: vbaof11.chm260002
f1_keywords:
- vbaof11.chm260002
ms.prod: office
api_name:
- Office.UserPermission.Permission
ms.assetid: 6350051e-c87f-f44d-2347-eab10775683b
ms.date: 06/08/2017
---


# UserPermission.Permission Property (Office)

 Returns or sets a **MsoPermission** constant as a **Long** value representing the permissions on the active document assigned to the user associated with the specified **UserPermission** object. Read/write.


## Syntax

 _expression_. **Permission**

 _expression_ A variable that represents a **UserPermission** object.


## Remarks

The  **Permission** property can be one or a combination of **MsoPermission** constants.

The  **UserPermission** object associates a set of permissions on the active document with a single user and an optional expiration date. The **Permission** property returns the set of user permissions determined by the specified **UserPermission** object. While some permissions granted through the user interface (such as **msoPermissionPrint** ) apply to all users, you can use the **UserPermission** object to assign them on a per-user basis with per-user expiration dates.


- The  **msoPermissionView** or **msoPermissionRead** option corresponds to the **Read** option in the user interface.
    
- The  **msoPermissionExtract** option corresponds to the **Allow users with read access to copy content** option in the user interface.
    
- The  **msoPermissionChange** option corresponds to the **Change** option in the user interface. The **msoPermissionChange** option represents the sum of **msoPermissionView** + **msoPermissionEdit** + **msoPermissionSave** + **msoPermissionExtract** and allows users to view, edit, copy, and save, but not print the document.
    
- The  **msoPermissionPrint** option corresponds to the **Print content** option in the user interface.
    
- The  **msoPermissionObjectModel** option corresponds to the **Access content programmatically** option in the user interface and allows users to access the document programmatically through its object model. Users without **msoPermissionObjectModel** cannot use the object model to determine their own rights, since programmatic access is disabled.
    

## Example

The following example uses the bitwise  **And** operator with the **Permission** property and an **msoPermission** constant to determine whether the second user has permission to save the active document.


```
 Dim irmPermission As Office.Permission 
 Dim irmUserPerm As Office.UserPermission 
 Set irmPermission = ActiveWorkbook.Permission 
 Set irmUserPerm = irmPermission.Item(2) 
 If irmUserPerm.Permission And Office.msoPermissionSave Then 
 MsgBox "User " &amp; irmUserPerm.UserId &amp; _ 
 " has permission to save this document.", _ 
 vbInformation + vbOKOnly, "IRM Information" 
 Else 
 MsgBox "User " &amp; irmUserPerm.UserId &amp; _ 
 " does NOT have permission to save this document.", _ 
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

