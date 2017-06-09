---
title: Permission.PolicyDescription Property (Office)
keywords: vbaof11.chm261011
f1_keywords:
- vbaof11.chm261011
ms.prod: office
api_name:
- Office.Permission.PolicyDescription
ms.assetid: 1ca10f9f-f03a-3a3f-2c12-21831a092f23
ms.date: 06/08/2017
---


# Permission.PolicyDescription Property (Office)

Gets the description of the permissions policy applied to the active document. Read-only.


## Syntax

 _expression_. **PolicyDescription**

 _expression_ A variable that represents a **Permission** object.


## Remarks

Microsoft Office Information Rights Management supports the use of administrative permission policies which list users and groups and their document permissions. The  **PolicyDescription** property returns the description of the policy applied to the active document, or a default value if a policy was not used.


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

