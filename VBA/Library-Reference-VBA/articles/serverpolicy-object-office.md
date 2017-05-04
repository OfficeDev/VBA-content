---
title: ServerPolicy Object (Office)
keywords: vbaof11.chm278010
f1_keywords:
- vbaof11.chm278010
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.ServerPolicy
ms.assetid: ce2a63d2-5deb-b94b-45d7-ed84e9be7deb
---


# ServerPolicy Object (Office)

Represents a policy specified for a document type stored on a server running Microsoft Office SharePoint Server.


## Remarks

The  **ServerPolicy** object is composed of individual **PolicyItem** objects representing the individual policy definitions for the active document.


## Example

The following example lists the name and description of all of the policy items for the active document.


```vb
Sub ListPolicyItems() 
Dim objSrvPolicy As ServerPolicy 
Dim objPolicyItem As PolicyItem 
Dim strPolicyItemList As String 
 
Set objSrvPolicy = ActiveDocument.ServerPolicy 
 
For Each objPolicyItem In objSrvPolicy 
 strPolicyItemList = "Policy Item " &; objPolicyItem.Name &; " - " &; _ 
 objPolicyItem.Description &; vbCrLf 
Next 
MsgBox (strPolicyItemList) 
 
End Sub
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

