---
title: Attachments.Remove Method (Outlook)
keywords: vbaol11.chm177
f1_keywords:
- vbaol11.chm177
ms.prod: outlook
api_name:
- Outlook.Attachments.Remove
ms.assetid: be49c973-b64e-84d9-1bf6-73b27a7e84f0
ms.date: 06/08/2017
---


# Attachments.Remove Method (Outlook)

Removes an object from the collection.


## Syntax

 _expression_ . **Remove** **_Index_**

 _expression_ A variable that represents an **Attachments** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The 1-based index value of the object within the collection.|

## Example

This Visual Basic for Applications (VBA) example uses the  **Remove** method to remove all attachments from a forwarded message before sending it on to Dan Wilson. Before running this example, replace 'Dan Wilson' with a valid recipient name.


```vb
Sub RemoveAttachmentBeforeForwarding() 
 
 Dim myinspector As Outlook.Inspector 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myattachments As Outlook.Attachments 
 
 
 
 Set myinspector = Application.ActiveInspector 
 
 If Not TypeName(myinspector) = "Nothing" Then 
 
 Set myItem = myinspector.CurrentItem.Forward 
 
 Set myattachments = myItem.Attachments 
 
 While myattachments.Count > 0 
 
 myattachments.Remove 1 
 
 Wend 
 
 myItem.Display 
 
 myItem.Recipients.Add "Dan Wilson" 
 
 myItem.Send 
 
 Else 
 
 MsgBox "There is no active inspector." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Attachments Object](attachments-object-outlook.md)

