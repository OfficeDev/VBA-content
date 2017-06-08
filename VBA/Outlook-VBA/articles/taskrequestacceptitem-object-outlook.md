---
title: TaskRequestAcceptItem Object (Outlook)
keywords: vbaol11.chm3008
f1_keywords:
- vbaol11.chm3008
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem
ms.assetid: a2905f72-0a67-b07d-7f85-84fe4de17c25
ms.date: 06/08/2017
---


# TaskRequestAcceptItem Object (Outlook)

Represents a response to a  **[TaskRequestItem](taskrequestitem-object-outlook.md)** sent by the initiating user.


## Remarks

If the delegated user accepts the task, the  **[ResponseState](taskitem-responsestate-property-outlook.md)** property is set to **olTaskAccept**. The associated **[TaskItem](taskitem-object-outlook.md)** is received by the delegator as a **TaskRequestAcceptItem** object.

Unlike other Microsoft Outlook objects, you cannot create this object.

Use the  **[GetAssociatedTask](taskrequestacceptitem-getassociatedtask-method-outlook.md)** method to return the **TaskItem** object that is associated with this **TaskRequestAcceptItem**. Work directly with the **TaskItem** object.


## Events



|**Name**|
|:-----|
|[AfterWrite](taskrequestacceptitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](taskrequestacceptitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](taskrequestacceptitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](taskrequestacceptitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](taskrequestacceptitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](taskrequestacceptitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](taskrequestacceptitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](taskrequestacceptitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](taskrequestacceptitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](taskrequestacceptitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](taskrequestacceptitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](taskrequestacceptitem-beforedelete-event-outlook.md)|
|[BeforeRead](taskrequestacceptitem-beforeread-event-outlook.md)|
|[Close](taskrequestacceptitem-close-event-outlook.md)|
|[CustomAction](taskrequestacceptitem-customaction-event-outlook.md)|
|[CustomPropertyChange](taskrequestacceptitem-custompropertychange-event-outlook.md)|
|[Forward](taskrequestacceptitem-forward-event-outlook.md)|
|[Open](taskrequestacceptitem-open-event-outlook.md)|
|[PropertyChange](taskrequestacceptitem-propertychange-event-outlook.md)|
|[Read](taskrequestacceptitem-read-event-outlook.md)|
|[ReadComplete](taskrequestacceptitem-readcomplete-event-outlook.md)|
|[Reply](taskrequestacceptitem-reply-event-outlook.md)|
|[ReplyAll](taskrequestacceptitem-replyall-event-outlook.md)|
|[Send](taskrequestacceptitem-send-event-outlook.md)|
|[Unload](taskrequestacceptitem-unload-event-outlook.md)|
|[Write](taskrequestacceptitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Close](taskrequestacceptitem-close-method-outlook.md)|
|[Copy](taskrequestacceptitem-copy-method-outlook.md)|
|[Delete](taskrequestacceptitem-delete-method-outlook.md)|
|[Display](taskrequestacceptitem-display-method-outlook.md)|
|[GetAssociatedTask](taskrequestacceptitem-getassociatedtask-method-outlook.md)|
|[GetConversation](taskrequestacceptitem-getconversation-method-outlook.md)|
|[Move](taskrequestacceptitem-move-method-outlook.md)|
|[PrintOut](taskrequestacceptitem-printout-method-outlook.md)|
|[Save](taskrequestacceptitem-save-method-outlook.md)|
|[SaveAs](taskrequestacceptitem-saveas-method-outlook.md)|
|[ShowCategoriesDialog](taskrequestacceptitem-showcategoriesdialog-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](taskrequestacceptitem-actions-property-outlook.md)|
|[Application](taskrequestacceptitem-application-property-outlook.md)|
|[Attachments](taskrequestacceptitem-attachments-property-outlook.md)|
|[AutoResolvedWinner](taskrequestacceptitem-autoresolvedwinner-property-outlook.md)|
|[BillingInformation](taskrequestacceptitem-billinginformation-property-outlook.md)|
|[Body](taskrequestacceptitem-body-property-outlook.md)|
|[Categories](taskrequestacceptitem-categories-property-outlook.md)|
|[Class](taskrequestacceptitem-class-property-outlook.md)|
|[Companies](taskrequestacceptitem-companies-property-outlook.md)|
|[Conflicts](taskrequestacceptitem-conflicts-property-outlook.md)|
|[ConversationID](taskrequestacceptitem-conversationid-property-outlook.md)|
|[ConversationIndex](taskrequestacceptitem-conversationindex-property-outlook.md)|
|[ConversationTopic](taskrequestacceptitem-conversationtopic-property-outlook.md)|
|[CreationTime](taskrequestacceptitem-creationtime-property-outlook.md)|
|[DownloadState](taskrequestacceptitem-downloadstate-property-outlook.md)|
|[EntryID](taskrequestacceptitem-entryid-property-outlook.md)|
|[FormDescription](taskrequestacceptitem-formdescription-property-outlook.md)|
|[GetInspector](taskrequestacceptitem-getinspector-property-outlook.md)|
|[Importance](taskrequestacceptitem-importance-property-outlook.md)|
|[IsConflict](taskrequestacceptitem-isconflict-property-outlook.md)|
|[ItemProperties](taskrequestacceptitem-itemproperties-property-outlook.md)|
|[LastModificationTime](taskrequestacceptitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](taskrequestacceptitem-markfordownload-property-outlook.md)|
|[MessageClass](taskrequestacceptitem-messageclass-property-outlook.md)|
|[Mileage](taskrequestacceptitem-mileage-property-outlook.md)|
|[NoAging](taskrequestacceptitem-noaging-property-outlook.md)|
|[OutlookInternalVersion](taskrequestacceptitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](taskrequestacceptitem-outlookversion-property-outlook.md)|
|[Parent](taskrequestacceptitem-parent-property-outlook.md)|
|[PropertyAccessor](taskrequestacceptitem-propertyaccessor-property-outlook.md)|
|[RTFBody](taskrequestacceptitem-rtfbody-property-outlook.md)|
|[Saved](taskrequestacceptitem-saved-property-outlook.md)|
|[Sensitivity](taskrequestacceptitem-sensitivity-property-outlook.md)|
|[Session](taskrequestacceptitem-session-property-outlook.md)|
|[Size](taskrequestacceptitem-size-property-outlook.md)|
|[Subject](taskrequestacceptitem-subject-property-outlook.md)|
|[UnRead](taskrequestacceptitem-unread-property-outlook.md)|
|[UserProperties](taskrequestacceptitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
