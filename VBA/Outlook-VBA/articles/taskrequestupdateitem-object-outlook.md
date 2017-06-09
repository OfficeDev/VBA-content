---
title: TaskRequestUpdateItem Object (Outlook)
keywords: vbaol11.chm3011
f1_keywords:
- vbaol11.chm3011
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem
ms.assetid: 5bc407fe-b3f6-3e46-8b91-e2ed96292cec
ms.date: 06/08/2017
---


# TaskRequestUpdateItem Object (Outlook)

Represents a response to a  **[TaskRequestItem](taskrequestitem-object-outlook.md)** sent by the initiating user.


## Remarks

If the delegated user updates the task by changing properties such as the  **[DueDate](taskitem-duedate-property-outlook.md)** or the **[Status](taskitem-status-property-outlook.md)**, and then sends it, the associated **[TaskItem](taskitem-object-outlook.md)** is received by the delegator as a **TaskRequestUpdateItem** object.

Unlike other Microsoft Outlook objects, you cannot create this object.

Use the  **[GetAssociatedTask](taskrequestupdateitem-getassociatedtask-method-outlook.md)** method to return the **TaskItem** object that is associated with this **TaskRequestUpdateItem**. Work directly with the **TaskItem** object


## Events



|**Name**|
|:-----|
|[AfterWrite](taskrequestupdateitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](taskrequestupdateitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](taskrequestupdateitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](taskrequestupdateitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](taskrequestupdateitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](taskrequestupdateitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](taskrequestupdateitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](taskrequestupdateitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](taskrequestupdateitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](taskrequestupdateitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](taskrequestupdateitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](taskrequestupdateitem-beforedelete-event-outlook.md)|
|[BeforeRead](taskrequestupdateitem-beforeread-event-outlook.md)|
|[Close](taskrequestupdateitem-close-event-outlook.md)|
|[CustomAction](taskrequestupdateitem-customaction-event-outlook.md)|
|[CustomPropertyChange](taskrequestupdateitem-custompropertychange-event-outlook.md)|
|[Forward](taskrequestupdateitem-forward-event-outlook.md)|
|[Open](taskrequestupdateitem-open-event-outlook.md)|
|[PropertyChange](taskrequestupdateitem-propertychange-event-outlook.md)|
|[Read](taskrequestupdateitem-read-event-outlook.md)|
|[ReadComplete](taskrequestupdateitem-readcomplete-event-outlook.md)|
|[Reply](taskrequestupdateitem-reply-event-outlook.md)|
|[ReplyAll](taskrequestupdateitem-replyall-event-outlook.md)|
|[Send](taskrequestupdateitem-send-event-outlook.md)|
|[Unload](taskrequestupdateitem-unload-event-outlook.md)|
|[Write](taskrequestupdateitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Close](taskrequestupdateitem-close-method-outlook.md)|
|[Copy](taskrequestupdateitem-copy-method-outlook.md)|
|[Delete](taskrequestupdateitem-delete-method-outlook.md)|
|[Display](taskrequestupdateitem-display-method-outlook.md)|
|[GetAssociatedTask](taskrequestupdateitem-getassociatedtask-method-outlook.md)|
|[GetConversation](taskrequestupdateitem-getconversation-method-outlook.md)|
|[Move](taskrequestupdateitem-move-method-outlook.md)|
|[PrintOut](taskrequestupdateitem-printout-method-outlook.md)|
|[Save](taskrequestupdateitem-save-method-outlook.md)|
|[SaveAs](taskrequestupdateitem-saveas-method-outlook.md)|
|[ShowCategoriesDialog](taskrequestupdateitem-showcategoriesdialog-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](taskrequestupdateitem-actions-property-outlook.md)|
|[Application](taskrequestupdateitem-application-property-outlook.md)|
|[Attachments](taskrequestupdateitem-attachments-property-outlook.md)|
|[AutoResolvedWinner](taskrequestupdateitem-autoresolvedwinner-property-outlook.md)|
|[BillingInformation](taskrequestupdateitem-billinginformation-property-outlook.md)|
|[Body](taskrequestupdateitem-body-property-outlook.md)|
|[Categories](taskrequestupdateitem-categories-property-outlook.md)|
|[Class](taskrequestupdateitem-class-property-outlook.md)|
|[Companies](taskrequestupdateitem-companies-property-outlook.md)|
|[Conflicts](taskrequestupdateitem-conflicts-property-outlook.md)|
|[ConversationID](taskrequestupdateitem-conversationid-property-outlook.md)|
|[ConversationIndex](taskrequestupdateitem-conversationindex-property-outlook.md)|
|[ConversationTopic](taskrequestupdateitem-conversationtopic-property-outlook.md)|
|[CreationTime](taskrequestupdateitem-creationtime-property-outlook.md)|
|[DownloadState](taskrequestupdateitem-downloadstate-property-outlook.md)|
|[EntryID](taskrequestupdateitem-entryid-property-outlook.md)|
|[FormDescription](taskrequestupdateitem-formdescription-property-outlook.md)|
|[GetInspector](taskrequestupdateitem-getinspector-property-outlook.md)|
|[Importance](taskrequestupdateitem-importance-property-outlook.md)|
|[IsConflict](taskrequestupdateitem-isconflict-property-outlook.md)|
|[ItemProperties](taskrequestupdateitem-itemproperties-property-outlook.md)|
|[LastModificationTime](taskrequestupdateitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](taskrequestupdateitem-markfordownload-property-outlook.md)|
|[MessageClass](taskrequestupdateitem-messageclass-property-outlook.md)|
|[Mileage](taskrequestupdateitem-mileage-property-outlook.md)|
|[NoAging](taskrequestupdateitem-noaging-property-outlook.md)|
|[OutlookInternalVersion](taskrequestupdateitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](taskrequestupdateitem-outlookversion-property-outlook.md)|
|[Parent](taskrequestupdateitem-parent-property-outlook.md)|
|[PropertyAccessor](taskrequestupdateitem-propertyaccessor-property-outlook.md)|
|[RTFBody](taskrequestupdateitem-rtfbody-property-outlook.md)|
|[Saved](taskrequestupdateitem-saved-property-outlook.md)|
|[Sensitivity](taskrequestupdateitem-sensitivity-property-outlook.md)|
|[Session](taskrequestupdateitem-session-property-outlook.md)|
|[Size](taskrequestupdateitem-size-property-outlook.md)|
|[Subject](taskrequestupdateitem-subject-property-outlook.md)|
|[UnRead](taskrequestupdateitem-unread-property-outlook.md)|
|[UserProperties](taskrequestupdateitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
