---
title: PolicyItem Object (Office)
keywords: vbaof11.chm278020
f1_keywords:
- vbaof11.chm278020
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.PolicyItem
ms.assetid: aced7bdc-8ef7-2621-f188-f3c1d44ab6dc
---


# PolicyItem Object (Office)

Represents an item within a  **ServerPolicy** object that contains the settings for one policy.


## Remarks

A policy item cannot exist outside the scope of a policy. Policy items are distinct conditions defined for a document stored on a server running Microsoft Office SharePoint Server.


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

