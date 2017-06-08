---
title: Use Information Rights Management for Visio
ms.prod: visio
ms.assetid: 3912bf98-3669-4de1-958a-f2fa7ec5cdad
ms.date: 06/08/2017
---


# Use Information Rights Management for Visio
Learn to use Information Rights Management (IRM) with Visio documents.




## Overview

You can use IRM support for Visio to programmatically:


- Protect a Visio document from unauthorized access.
    
- Manage individual and group access to an IRM-protected Visio document.
    
- Change IRM permissions for a Visio document.
    

## Requirements

To use IRM for Visio, you must install the following: 


- Visio
    
-  [Windows Rights Management Client](http://www.microsoft.com/en-us/download/details.aspx?id=4909)
    

## Protecting a Visio document

To protect the active document, use the  [Permission.Add](http://msdn.microsoft.com/en-us/library/office/ff863139%28v=office.15%29.aspx) method. To check whether a document is protected, use the **Enabled** property of the **Permission** object.

To remove protection from the active document, use the  [Permission.RemoveAll](http://msdn.microsoft.com/en-us/library/office/ff861135%28v=office.15%29.aspx) method, or use the [UserPermission.Remove](http://msdn.microsoft.com/en-us/library/office/ff864865%28v=office.15%29.aspx) method for each user that has access.


## Managing user access to an IRM-protected document

To give permissions on the active document to a specified user, use the  **Permission.Add** method. The **Permission** property can be one or a combination of **msoPermission** constants from the following table.


****


|**msoPermission constant**|**Meaning**|
|:-----|:-----|
| **msoPermissionView**|Read access|
| **msoPermissionRead**|Read access|
| **msoPermissionEdit**|Edit access|
| **msoPermission Save**|Save access|
| **msoPermissionExtract**|Copy access, if the user also has read access|
| **msoPermissionChange**|Access to view, edit, copy, and save but not to print the document. This is equivalent to  **msoPermissionView** + **msoPermissionEdit** + **msoPermissionSave** + **msoPermissionExtract**.|
| **msoPermissionPrint**|Print access|
| **msoPermissionObjModel**|A user can access the document programmatically. All users need this permission to work with a protected document or to check their permissions on the document.|
| **msoPermissionFullControl**|Full control over the document. All permissions are enabled.|
To check permissions for a user, use the  [UserPermission.Permission](http://msdn.microsoft.com/en-us/library/office/ff862094%28v=office.15%29.aspx) property.

To apply permissions by using an administrative policy, use the  [Permission.ApplyPolicy](http://msdn.microsoft.com/en-us/library/office/ff864678%28v=office.15%29.aspx) method. Use the **PermissionFromPolicy**,  **PolicyName**, and  **PolicyDescription** properties to return policy information.

To remove permissions for a specified user, use the  [UserPermission.Remove](http://msdn.microsoft.com/en-us/library/office/ff864865%28v=office.15%29.aspx) method. To remove all restrictions on the active document, use the [Permission.RemoveAll](http://msdn.microsoft.com/en-us/library/office/ff861135%28v=office.15%29.aspx) method.


## Accessing an IRM-protected document

To access a protected document, a user needs the following:


-  **msoPermissionObjModel**
    
- The specific IRM permissions for any method or event that is used.
    
The following tables show the required permissions. Use the corresponding  **msoPermission** constants from the previous section. For almost all APIs, the user needs the Edit ( **msoPermissionEdit**) permission. Note that if the user has Full Control ( **msoPermissionFullControl**), all APIs are enabled.


**IRM permissions for Methods**


|**Method**|**Required permission**|
|:-----|:-----|
| **Copy**|Read and Copy|
| **Copy** ( **Selection** object)|Read and Copy|
| **Copy** ( **Shape** object)|Read and Copy|
| **GetFilterCommands**|Read|
| **GetFilterObjects**|Read|
| **GetFilterSRC**|Read|
| **GetFormulas[U]**|Read|
| **GetNames[U]**|Read|
| **GetPolylineData**|Read|
| **GetResults**|Read|
| **GetViewRect**|Read|
| **Open** ( **Documents** collection)|Read|
| **OpenEx**|Read|
| **Print**|Read and Print|
| **PrintTile**|Read and Print|
|For all other events not listed here|Edit|

**IRM permissions for Events**


|**Event**|**Required permission**|
|:-----|:-----|
|DocumentOpened|Read|
|WindowOpened|Read|
|WindowTurnedToPage|Read|
|For all other events not listed here|Edit|

## Examples




### Apply a user's permission example

This example gives a user a combination of Read and Edit permissions on the active document and sets an expiration date for these permissions.


```vb
Dim objPermission As Office.Permission
Dim objUserPerm As Office.UserPermission

Set objPermission = ActiveDocument.Permission
Set objUserPerm = objPermission.Add( _
"<user>@<domain>.com", _
msoPermissionRead + msoPermissionEdit, #12/31/2016#)
MsgBox "Permissions added for " &; _
objUserPerm.UserId, _
vbInformation + vbOKOnly, _
"Permissions Added"
Set objUserPerm = Nothing

```


### Apply administrative permission policy example

This example checks whether the active document is protected and, if it is not, protects the document and applies an administrative permission policy.


```vb
Dim irmPermission As Office.Permission 
 Set irmPermission = ActiveDocument.Permission 
 Dim strIRMInfo As String 
 Select Case irmPermission.Enabled 
 Case True 
 strIRMInfo = "Permissions are already restricted on this document." 
 Case False 
 With irmPermission 
 .Enabled = True 
 .ApplyPolicy ("\\server\share\permissionpolicy.xml") 
 End With 
 strIRMInfo = "Permissions are now restricted on this document " &; _ 
 vbCrLf &; _ 
 " and the permission policy has been applied." 
 End Select 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing
```


### List permitted users example

This example checks whether the active document is protected and, if it is, lists users and their assigned permissions.


```vb
Dim irmPermission As Office.Permission 
 Dim irmUserPerm As Office.UserPermission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveDocument.Permission 
 If irmPermission.Enabled Then 
 For Each irmUserPerm In irmPermission 
 strIRMInfo = strIRMInfo &; irmUserPerm.UserId &; vbCrLf &; _ 
 " - Permissions: " &; irmUserPerm.Permission &; vbCrLf &; _ 
 " - Expiration Date: " &; irmUserPerm.ExpirationDate &; vbCrLf 
 Next 
 MsgBox strIRMInfo, _ 
 vbInformation + vbOKOnly, "IRM Information" 
 Else 
 MsgBox "This document is not restricted.", _ 
 vbInformation + vbOKOnly, "IRM Information" 
 End If 
 Set irmUserPerm = Nothing 
 Set irmPermission = Nothing
```


### Remove a user's permissions example

This example removes the second user's permissions on the active document from the document's  **Permission** collection.


```vb
Dim irmPermission As Office.Permission 
 Dim irmUserPerm As Office.UserPermission 
 Set irmPermission = ActiveDocument.Permission 
 Set irmUserPerm = irmPermission.Item(2) 
 irmUserPerm.Remove 
 MsgBox "Permission removed.", _ 
 vbInformation + vbOKOnly, "IRM Information" 
 Set irmUserPerm = Nothing 
 Set irmPermission = Nothing
```


