---
title: Permission.PermissionFromPolicy Property (Office)
keywords: vbaof11.chm261014
f1_keywords:
- vbaof11.chm261014
ms.prod: office
api_name:
- Office.Permission.PermissionFromPolicy
ms.assetid: aa6be9a8-a351-f9bb-99f8-a547583f2e62
ms.date: 06/08/2017
---


# Permission.PermissionFromPolicy Property (Office)

Gets a  **Boolean** value that indicates whether a permission policy has been applied to the active document. Read-only.


## Syntax

 _expression_. **PermissionFromPolicy**

 _expression_ A variable that represents a **Permission** object.


## Remarks

Information Rights Management in Microsoft Office supports the use of administrative permission policies which list users and groups and their document permissions. The  **PermissionFromPolicy** property returns a **Boolean** value that indicates whether a permission policy was applied to the active document the last time permissions were enabled on the document.

The  **PermissionFromPolicy** property always returns **False** when checked by a non-owner of the document, even when the user has object model permissions.


## Example

The following example displays permission policy information about the active document.


```
 Dim irmPermission As Office.Permission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 strIRMInfo = "Permissions are restricted on this document." &amp; vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo &amp; " Permissions applied from policy:" &amp; vbCrLf &amp; _ 
 " Policy name: " &amp; irmPermission.PolicyName &amp; vbCrLf &amp; _ 
 " Policy description: " &amp; irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo &amp; " Default permissions applied." &amp; vbCrLf &amp; _ 
 " Default policy name: " &amp; irmPermission.PolicyName &amp; vbCrLf &amp; _ 
 " Default policy description: " &amp; irmPermission.PolicyDescription 
 End If 
 Else 
 strIRMInfo = "Permission are NOT restricted on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing 

```


## See also


#### Concepts


[Permission Object](permission-object-office.md)
#### Other resources


[Permission Object Members](permission-members-office.md)

