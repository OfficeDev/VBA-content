---
title: TaskRequestItem Object (Outlook)
keywords: vbaol11.chm3010
f1_keywords:
- vbaol11.chm3010
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem
ms.assetid: 2908a28a-634c-e786-aa53-f3e32038b727
ms.date: 06/08/2017
---


# TaskRequestItem Object (Outlook)

Represents a change to the recipient's Tasks list initiated by another party or as a result of a group tasking.


## Remarks

Unlike other Microsoft Outlook objects, you cannot create this object. When the sender applies the  **[Assign](taskitem-assign-method-outlook.md)** and **[Send](taskitem-send-method-outlook.md)** methods to a **[TaskItem](taskitem-object-outlook.md)** object to assign (delegate) the associated task to another user, the **TaskRequestItem** object is created when the item is received in the recipient's Inbox.

Use the  **[GetAssociatedTask](taskrequestitem-getassociatedtask-method-outlook.md)** method to return the **TaskItem** object, and work directly with the **TaskItem** object to respond to the request.


## Example

The following Visual Basic for Applications (VBA) example creates a simple task, assigns it to another user, and sends it. When the task request arrives in the recipient's Inbox, it is received as a  **TaskRequestItem**.






```
Sub SendTask() 
 
 Dim myItem As Outlook.TaskItem 
 
 Dim myDelegate As Outlook.Recipient 
 
 
 
 Set myItem = Application.CreateItem(olTaskItem) 
 
 myItem.Assign 
 
 Set myDelegate = myItem.Recipients.Add("Jeff Smith") 
 
 myItem.Subject = "Prepare Agenda For Meeting" 
 
 myItem.DueDate = #9/20/97# 
 
 myItem.Send 
 
End Sub
```


## Events



|**Name**|
|:-----|
|[AfterWrite](taskrequestitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](taskrequestitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](taskrequestitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](taskrequestitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](taskrequestitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](taskrequestitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](taskrequestitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](taskrequestitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](taskrequestitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](taskrequestitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](taskrequestitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](taskrequestitem-beforedelete-event-outlook.md)|
|[BeforeRead](taskrequestitem-beforeread-event-outlook.md)|
|[Close](taskrequestitem-close-event-outlook.md)|
|[CustomAction](taskrequestitem-customaction-event-outlook.md)|
|[CustomPropertyChange](taskrequestitem-custompropertychange-event-outlook.md)|
|[Forward](taskrequestitem-forward-event-outlook.md)|
|[Open](taskrequestitem-open-event-outlook.md)|
|[PropertyChange](taskrequestitem-propertychange-event-outlook.md)|
|[Read](taskrequestitem-read-event-outlook.md)|
|[ReadComplete](taskrequestitem-readcomplete-event-outlook.md)|
|[Reply](taskrequestitem-reply-event-outlook.md)|
|[ReplyAll](taskrequestitem-replyall-event-outlook.md)|
|[Send](taskrequestitem-send-event-outlook.md)|
|[Unload](taskrequestitem-unload-event-outlook.md)|
|[Write](taskrequestitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Close](taskrequestitem-close-method-outlook.md)|
|[Copy](taskrequestitem-copy-method-outlook.md)|
|[Delete](taskrequestitem-delete-method-outlook.md)|
|[Display](taskrequestitem-display-method-outlook.md)|
|[GetAssociatedTask](taskrequestitem-getassociatedtask-method-outlook.md)|
|[GetConversation](taskrequestitem-getconversation-method-outlook.md)|
|[Move](taskrequestitem-move-method-outlook.md)|
|[PrintOut](taskrequestitem-printout-method-outlook.md)|
|[Save](taskrequestitem-save-method-outlook.md)|
|[SaveAs](taskrequestitem-saveas-method-outlook.md)|
|[ShowCategoriesDialog](taskrequestitem-showcategoriesdialog-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](taskrequestitem-actions-property-outlook.md)|
|[Application](taskrequestitem-application-property-outlook.md)|
|[Attachments](taskrequestitem-attachments-property-outlook.md)|
|[AutoResolvedWinner](taskrequestitem-autoresolvedwinner-property-outlook.md)|
|[BillingInformation](taskrequestitem-billinginformation-property-outlook.md)|
|[Body](taskrequestitem-body-property-outlook.md)|
|[Categories](taskrequestitem-categories-property-outlook.md)|
|[Class](taskrequestitem-class-property-outlook.md)|
|[Companies](taskrequestitem-companies-property-outlook.md)|
|[Conflicts](taskrequestitem-conflicts-property-outlook.md)|
|[ConversationID](taskrequestitem-conversationid-property-outlook.md)|
|[ConversationIndex](taskrequestitem-conversationindex-property-outlook.md)|
|[ConversationTopic](taskrequestitem-conversationtopic-property-outlook.md)|
|[CreationTime](taskrequestitem-creationtime-property-outlook.md)|
|[DownloadState](taskrequestitem-downloadstate-property-outlook.md)|
|[EntryID](taskrequestitem-entryid-property-outlook.md)|
|[FormDescription](taskrequestitem-formdescription-property-outlook.md)|
|[GetInspector](taskrequestitem-getinspector-property-outlook.md)|
|[Importance](taskrequestitem-importance-property-outlook.md)|
|[IsConflict](taskrequestitem-isconflict-property-outlook.md)|
|[ItemProperties](taskrequestitem-itemproperties-property-outlook.md)|
|[LastModificationTime](taskrequestitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](taskrequestitem-markfordownload-property-outlook.md)|
|[MessageClass](taskrequestitem-messageclass-property-outlook.md)|
|[Mileage](taskrequestitem-mileage-property-outlook.md)|
|[NoAging](taskrequestitem-noaging-property-outlook.md)|
|[OutlookInternalVersion](taskrequestitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](taskrequestitem-outlookversion-property-outlook.md)|
|[Parent](taskrequestitem-parent-property-outlook.md)|
|[PropertyAccessor](taskrequestitem-propertyaccessor-property-outlook.md)|
|[RTFBody](taskrequestitem-rtfbody-property-outlook.md)|
|[Saved](taskrequestitem-saved-property-outlook.md)|
|[Sensitivity](taskrequestitem-sensitivity-property-outlook.md)|
|[Session](taskrequestitem-session-property-outlook.md)|
|[Size](taskrequestitem-size-property-outlook.md)|
|[Subject](taskrequestitem-subject-property-outlook.md)|
|[UnRead](taskrequestitem-unread-property-outlook.md)|
|[UserProperties](taskrequestitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
