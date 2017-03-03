---
title: Permission.PolicyName Property (Office)
keywords: vbaof11.chm261010
f1_keywords:
- vbaof11.chm261010
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.Permission.PolicyName
ms.assetid: 2a76eac3-5012-6c6c-ab5a-388151f50e27
---


# Permission.PolicyName Property (Office)

Gets the name of the permissions policy applied to the active document. Read-only.


## Syntax

 _expression_. **PolicyName**

 _expression_ A variable that represents a **Permission** object.


## Remarks

Microsoft Office Information Rights Management supports the use of administrative permission policies which list users and groups and their document permissions. The  **PolicyName** property returns the name of the policy applied to the active document, or a default value if a policy was not used.


## Example

The following example displays permission policy information about the active document.


```vb
 Dim irmPermission As Office.Permission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 strIRMInfo = "Permissions are restricted on this document." &; vbCrLf 
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
 strIRMInfo = "Permission are NOT restricted on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing 

```


## See also


#### Concepts


[Permission Object](permission-object-office.md)

