---
title: Permission Object (Office)
keywords: vbaof11.chm261000
f1_keywords:
- vbaof11.chm261000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.Permission
ms.assetid: 4bdf7058-d4ba-0bd4-c5cd-141d67245ced
---


# Permission Object (Office)

The  **Permission** property of the **Document** object in Microsoft Word, a **Workbook** object in Microsoft Excel, and a **Presentation** object in Microsoft PowerPoint returns a **Permission** object.


## Remarks

Use the  **Permission** object to restrict permissions to the active document and to return or set specific permissions settings.

The  **Permission** object gives access to a collection of **UserPermission** objects. Use the **UserPermission** object to associate specific sets of rights with individual users. While some permissions granted through the user interface (such as **msoPermissionPrint** ) apply to all users, you can use the **UserPermission** object to assign them on a per-user basis with per-user expiration dates.

Microsoft Office Information Rights Management supports the use of administrative permission policies which list users and groups and their document permissions. Use the  **ApplyPolicy** method to apply a permission policy, and the **PermissionFromPolicy**, **PolicyName**, and **PolicyDescription** properties to return policy information.

The  **Permission** object model is available whether permissions are restricted on the active document or not . The **Permission** property of the **Document**, **Workbook**, and **Presentation** objects does not return **Nothing** when the active document does not have restricted permissions. Use the **Enabled** property to determine whether a document has restricted permissions.

Use of the  **Permission** object raises an error when the Windows Rights Management client is not installed.


## Example

The following example returns information about the permissions settings on the active document.


```vb
Dim irmPermission As Office.Permission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 strIRMInfo = "Permissions are restricted on this document." &; vbCrLf 
 strIRMInfo = strIRMInfo &; " View in trusted browser: " &; _ 
 irmPermission.EnableTrustedBrowser &; vbCrLf &; _ 
 " Document author: " &; irmPermission.DocumentAuthor &; vbCrLf &; _ 
 " Users with permissions: " &; irmPermission.Count &; vbCrLf &; _ 
 " Cache licenses: " &; irmPermission.StoreLicenses &; vbCrLf &; _ 
 " Request permission URL: " &; irmPermission.RequestPermissionURL &; vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo &; " Permissions applied from policy:" &; vbCrLf &; _ 
 " Policy name: " &; irmPermission.PolicyName &; vbCrLf &; _ 
 " Policy description: " &; irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo &; " Default permissions applied." &; vbCrLf &; _ 
 " Default policy name: " &; irmPermission.PolicyName &; vbCrLf &; _ 
 " Default policy description: " &; irmPermission.PolicyDescription 
 End If 
 Else 
 strIRMInfo = "Permissions are NOT restricted on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

