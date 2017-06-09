---
title: ServerPolicy Object (Office)
keywords: vbaof11.chm278010
f1_keywords:
- vbaof11.chm278010
ms.prod: office
api_name:
- Office.ServerPolicy
ms.assetid: ce2a63d2-5deb-b94b-45d7-ed84e9be7deb
ms.date: 06/08/2017
---


# ServerPolicy Object (Office)

Represents a policy specified for a document type stored on a server running Microsoft Office SharePoint Server.


## Remarks

The  **ServerPolicy** object is composed of individual **PolicyItem** objects representing the individual policy definitions for the active document.


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
|[Application](serverpolicy-application-property-office.md)|
|[BlockPreview](serverpolicy-blockpreview-property-office.md)|
|[Count](serverpolicy-count-property-office.md)|
|[Creator](serverpolicy-creator-property-office.md)|
|[Description](serverpolicy-description-property-office.md)|
|[Id](serverpolicy-id-property-office.md)|
|[Item](serverpolicy-item-property-office.md)|
|[Name](serverpolicy-name-property-office.md)|
|[Parent](serverpolicy-parent-property-office.md)|
|[Statement](serverpolicy-statement-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
