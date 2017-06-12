---
title: MailItem.AttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.AttachmentRead
ms.assetid: 9da23894-0867-aac8-2275-251e32ad4180
ms.date: 06/08/2017
---


# MailItem.AttachmentRead Event (Outlook)

Occurs when an attachment in an instance of the parent object has been opened for reading.


## Syntax

 _expression_ . **AttachmentRead**( **_Attachment_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** that was opened.|

## Example

This Visual Basic for Applications (VBA) example displays a message when the user tries to read an attachment. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `TestAttachRead()` procedure should be called before the event procedure can be called by Microsoft Outlook. For this example to run, there has to be at least one item in the Inbox with subject as 'Test' and containing at least one attachment.


```vb
Public WithEvents myItem As outlook.MailItem 
 
 
 
Private Sub myItem_AttachmentRead(ByVal myAttachment As Outlook.Attachment) 
 
 If myAttachment.Type = olByValue Then 
 
 MsgBox "If you change this file, also save your changes to the original file." 
 
 End If 
 
End Sub 
 
 
 
Public Sub TestAttachRead() 
 
 Dim atts As Outlook.Attachments 
 
 Dim myAttachment As Outlook.Attachment 
 
 
 
 Set myItem = Application.ActiveExplorer.CurrentFolder.Items("Test") 
 
 Set atts = myItem.Attachments 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

