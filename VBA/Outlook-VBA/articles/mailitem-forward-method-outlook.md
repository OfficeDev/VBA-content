---
title: MailItem.Forward Method (Outlook)
keywords: vbaol11.chm1366
f1_keywords:
- vbaol11.chm1366
ms.prod: outlook
api_name:
- Outlook.MailItem.Forward
ms.assetid: 5b8c2261-c5ac-fd80-8acf-dfa645a04a1e
ms.date: 06/08/2017
---


# MailItem.Forward Method (Outlook)

Executes the  **Forward** action for an item and returns the resulting copy as a **[MailItem](mailitem-object-outlook.md)** object.


## Syntax

 _expression_ . **Forward**

 _expression_ A variable that represents a **MailItem** object.


### Return Value

A  **MailItem** object that represents the new mail item.


## Example

This Visual Basic for Applications (VBA) example uses the  **[Remove](attachments-remove-method-outlook.md)** method to remove all attachments from a forwarded message before sending it on to Dan Wilson. To run this example, replace 'Dan Wilson' with a valid recipient name and keep a mail item that contains at least one attachment open in the active window.


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


[MailItem Object](mailitem-object-outlook.md)

