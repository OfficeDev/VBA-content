---
title: MailItem.Attachments Property (Outlook)
keywords: vbaol11.chm1295
f1_keywords:
- vbaol11.chm1295
ms.prod: outlook
api_name:
- Outlook.MailItem.Attachments
ms.assetid: 71f82397-00f3-5660-1211-ebf8b229fff3
ms.date: 06/08/2017
---


# MailItem.Attachments Property (Outlook)

Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.


## Syntax

 _expression_ . **Attachments**

 _expression_ A variable that represents a **MailItem** object.


## Example

This Visual Basic for Applications (VBA) example uses the  **[Attachments.Remove](attachments-remove-method-outlook.md)** method to remove all attachments from a forwarded mail message before sending it on to 'Dan Wilson'. To run this example, replace 'Dan Wilson' with a valid recipient's name and keep an item with attachments open in an inspector window.


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

