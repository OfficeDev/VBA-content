---
title: MailItem.AttachmentAdd Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.AttachmentAdd
ms.assetid: ae95c10b-f8dc-0341-4153-c7805d973df9
ms.date: 06/08/2017
---


# MailItem.AttachmentAdd Event (Outlook)

Occurs when an attachment has been added to an instance of the parent object.


## Syntax

 _expression_ . **AttachmentAdd**( **_Attachment_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** that was added to the item.|

## Example

This Visual Basic for Applications (VBA) example checks the size of the item after an attachment has been added and displays a warning if the size exceeds 500,000 bytes. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `TestAttachAdd()` procedure should be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents newItem As Outlook.MailItem 
 
 
 
Private Sub newItem_AttachmentAdd(ByVal newAttachment As Attachment) 
 
 If newAttachment.Type = olByValue Then 
 
 newItem.Save 
 
 If newItem.Size > 500000 Then 
 
 MsgBox "Warning: Item size is now " &; newItem.Size &; " bytes." 
 
 End If 
 
 End If 
 
End Sub 
 
 
 
Public Sub TestAttachAdd() 
 
 Dim atts As Outlook.Attachments 
 
 Dim newAttachment As Outlook.Attachment 
 
 
 
 Set newItem = Application.CreateItem(olMailItem) 
 
 newItem.Subject = "Test attachment" 
 
 Set atts = newItem.Attachments 
 
 Set newAttachment = atts.Add("C:\Test.txt", olByValue) 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

