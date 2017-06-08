---
title: Permission.RequestPermissionURL Property (Office)
keywords: vbaof11.chm261009
f1_keywords:
- vbaof11.chm261009
ms.prod: office
api_name:
- Office.Permission.RequestPermissionURL
ms.assetid: 7d37d706-a7bf-9cb0-8930-299bd2bf37b0
ms.date: 06/08/2017
---


# Permission.RequestPermissionURL Property (Office)

Gets or sets the file or Web site URL to visit or the e-mail address to contact for users who need additional permissions on the active document. Read/write.


## Syntax

 _expression_. **RequestPermissionURL**

 _expression_ A variable that represents a **Permission** object.


## Remarks

The ** RequestPermissionURL** setting corresponds to the **Users can request additional permissions from** option in the permissions user interface. Use the **RequestPermissionURL** property to specify a file, a Web site, or an e-mail contact from which users can request, or learn how to request, additional permissions on the active document, for example:


- A Web address:  `http://companyserver/request_permissions.asp`
    
- A file:  `\\companyserver\share\requesting_permissions.txt`
    
- An email address:  `mailto:permissionsmgr@example.com?Subject=Request%20permissions`
    

## Example

The following example displays information about the permissions settings of the active document, including the  **RequestPermissionURL** setting.


```
 Dim irmPermission As Office.Permission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 strIRMInfo = "Permissions are restricted on this document." &amp; vbCrLf 
 strIRMInfo = strIRMInfo &amp; " View in trusted browser: " &amp; _ 
 irmPermission.EnableTrustedBrowser &amp; vbCrLf &amp; _ 
 " Document author: " &amp; irmPermission.DocumentAuthor &amp; vbCrLf &amp; _ 
 " Users with permissions: " &amp; irmPermission.Count &amp; vbCrLf &amp; _ 
 " Cache licenses locally: " &amp; irmPermission.StoreLicenses &amp; vbCrLf &amp; _ 
 " Request permission URL: " &amp; irmPermission.RequestPermissionURL &amp; vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo &amp; " Permissions applied from policy:" &amp; vbCrLf &amp; _ 
 " Policy name: " &amp; irmPermission.PolicyName &amp; vbCrLf &amp; _ 
 " Policy description: " &amp; irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo &amp; " Default permissions applied." 
 End If 
 Else 
 strIRMInfo = "Permissions are NOT restricted on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing 

```


## See also


#### Concepts


[Permission Object](permission-object-office.md)
#### Other resources


[Permission Object Members](permission-members-office.md)

