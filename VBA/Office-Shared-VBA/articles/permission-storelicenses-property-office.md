---
title: Permission.StoreLicenses Property (Office)
keywords: vbaof11.chm261012
f1_keywords:
- vbaof11.chm261012
ms.prod: office
api_name:
- Office.Permission.StoreLicenses
ms.assetid: c08e088c-8cdf-baa0-56e4-3d4d6f3caab8
ms.date: 06/08/2017
---


# Permission.StoreLicenses Property (Office)

Gets or sets a  **Boolean** value that indicates whether the user's license to view the active document should be cached to allow offline viewing when the user cannot connect to a rights management server. Read/write.


## Syntax

 _expression_. **StoreLicenses**

 _expression_ A variable that represents a **Permission** object.


## Remarks

The  **StoreLicenses** property corresponds to (and its value is the opposite of) the **Require a connection to verify a user's permission** option in the permissions user interface. When **StoreLicenses** is **False**, users other than the document owner must connect to the rights management server and acquire the license to work with the document each time they open it when content is protected using the Information Rights Management service provided in Microsoft Office.


## Example

The following example displays information about the permissions settings of the active document, including the  **StoreLicenses** setting.


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
 strIRMInfo = strIRMInfo &amp; " Custom permissions applied." 
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

