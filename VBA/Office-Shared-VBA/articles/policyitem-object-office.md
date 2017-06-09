---
title: PolicyItem Object (Office)
keywords: vbaof11.chm278020
f1_keywords:
- vbaof11.chm278020
ms.prod: office
api_name:
- Office.PolicyItem
ms.assetid: aced7bdc-8ef7-2621-f188-f3c1d44ab6dc
ms.date: 06/08/2017
---


# PolicyItem Object (Office)

Represents an item within a  **ServerPolicy** object that contains the settings for one policy.


## Remarks

A policy item cannot exist outside the scope of a policy. Policy items are distinct conditions defined for a document stored on a server running Microsoft Office SharePoint Server.


## Example

The following example lists the name and description of all of the policy items for the active document.


```
Sub ListPolicyItems() 
Dim objSrvPolicy As ServerPolicy 
Dim objPolicyItem As PolicyItem 
Dim strPolicyItemList As String 
 
Set objSrvPolicy = ActiveDocument.ServerPolicy 
 
For Each objPolicyItem In objSrvPolicy 
 strPolicyItemList = "Policy Item " &amp; objPolicyItem.Name &amp; " - " &amp; _ 
 objPolicyItem.Description &amp; vbCrLf 
Next 
MsgBox (strPolicyItemList) 
 
End Sub 

```


## Properties



|**Name**|
|:-----|
|[Application](policyitem-application-property-office.md)|
|[Creator](policyitem-creator-property-office.md)|
|[Data](policyitem-data-property-office.md)|
|[Description](policyitem-description-property-office.md)|
|[Id](policyitem-id-property-office.md)|
|[Name](policyitem-name-property-office.md)|
|[Parent](policyitem-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
