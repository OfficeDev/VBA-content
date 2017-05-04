---
title: Permission.StoreLicenses Property (Office)
keywords: vbaof11.chm261012
f1_keywords:
- vbaof11.chm261012
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.Permission.StoreLicenses
ms.assetid: c08e088c-8cdf-baa0-56e4-3d4d6f3caab8
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
 " Cache licenses locally: " &; irmPermission.StoreLicenses &; vbCrLf &; _ 
 " Request permission URL: " &; irmPermission.RequestPermissionURL &; vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo &; " Permissions applied from policy:" &; vbCrLf &; _ 
 " Policy name: " &; irmPermission.PolicyName &; vbCrLf &; _ 
 " Policy description: " &; irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo &; " Custom permissions applied." 
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

