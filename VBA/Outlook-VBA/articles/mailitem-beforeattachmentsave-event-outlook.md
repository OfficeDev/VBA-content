---
title: MailItem.BeforeAttachmentSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.BeforeAttachmentSave
ms.assetid: b36eb8dc-3128-c75c-9c2d-b5321d93680c
ms.date: 06/08/2017
---


# MailItem.BeforeAttachmentSave Event (Outlook)

Occurs just before an attachment is saved.


## Syntax

 _expression_ . **BeforeAttachmentSave**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be saved.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the save operation is not completed and the attachment is not changed.|

## Remarks

This event corresponds to when attachments are saved to the messaging store. The  **BeforeAttachmentSave** event occurs just before an attachment is saved when an item is saved. If a user edits an attachment and then saves those changes, the **BeforeAttachmentSave** event will not occur at that time; instead it will occur when the item itself is later saved. It also does not occur when the attachment is saved on the hard disk using the **SaveAsFile** method.

In VBScript, if you set the return value of this function to  **False** , the save operation is cancelled and the attachment is not changed.


## Example

This Visual Basic for Applications (VBA) example notifies the user that the user is not allowed to save the attachment. The  _Cancel_ argument is set to **True** to cancel the save operation. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `TestAttachSave()` procedure should be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
Private Sub myItem_BeforeAttachmentSave(ByVal myAttachment As Attachment, Cancel As Boolean) 
 MsgBox "You are not allowed to save " &; myAttachment.FileName 
 Cancel = True 
End Sub 
 
Public Sub TestAttachSave() 
 Set myItem = Application.ActiveInspector.CurrentItem 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

