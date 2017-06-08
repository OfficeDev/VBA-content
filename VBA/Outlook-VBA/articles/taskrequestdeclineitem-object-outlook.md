---
title: TaskRequestDeclineItem Object (Outlook)
keywords: vbaol11.chm3009
f1_keywords:
- vbaol11.chm3009
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem
ms.assetid: e842c7c0-7943-9219-329b-30b892ab99b0
ms.date: 06/08/2017
---


# TaskRequestDeclineItem Object (Outlook)

Represents a response to a  **[TaskRequestItem](taskrequestitem-object-outlook.md)** sent by the initiating user.


## Remarks

If the delegated user declines the task, the  **[ResponseState](taskitem-responsestate-property-outlook.md)** property is set to **olTaskDecline**. The associated **[TaskItem](taskitem-object-outlook.md)** is received by the delegator as a **TaskRequestDeclineItem** object.

Unlike other Microsoft Outlook objects, you cannot create this object.

Use the  **[GetAssociatedTask](taskrequestdeclineitem-getassociatedtask-method-outlook.md)** method to return the **TaskItem** object that is associated with this **TaskRequestDeclineItem**. Work directly with the **TaskItem** object.


## Events



|**Name**|
|:-----|
|[AfterWrite](taskrequestdeclineitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](taskrequestdeclineitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](taskrequestdeclineitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](taskrequestdeclineitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](taskrequestdeclineitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](taskrequestdeclineitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](taskrequestdeclineitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](taskrequestdeclineitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](taskrequestdeclineitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](taskrequestdeclineitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](taskrequestdeclineitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](taskrequestdeclineitem-beforedelete-event-outlook.md)|
|[BeforeRead](taskrequestdeclineitem-beforeread-event-outlook.md)|
|[Close](taskrequestdeclineitem-close-event-outlook.md)|
|[CustomAction](taskrequestdeclineitem-customaction-event-outlook.md)|
|[CustomPropertyChange](taskrequestdeclineitem-custompropertychange-event-outlook.md)|
|[Forward](taskrequestdeclineitem-forward-event-outlook.md)|
|[Open](taskrequestdeclineitem-open-event-outlook.md)|
|[PropertyChange](taskrequestdeclineitem-propertychange-event-outlook.md)|
|[Read](taskrequestdeclineitem-read-event-outlook.md)|
|[ReadComplete](taskrequestdeclineitem-readcomplete-event-outlook.md)|
|[Reply](taskrequestdeclineitem-reply-event-outlook.md)|
|[ReplyAll](taskrequestdeclineitem-replyall-event-outlook.md)|
|[Send](taskrequestdeclineitem-send-event-outlook.md)|
|[Unload](taskrequestdeclineitem-unload-event-outlook.md)|
|[Write](taskrequestdeclineitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Close](taskrequestdeclineitem-close-method-outlook.md)|
|[Copy](taskrequestdeclineitem-copy-method-outlook.md)|
|[Delete](taskrequestdeclineitem-delete-method-outlook.md)|
|[Display](taskrequestdeclineitem-display-method-outlook.md)|
|[GetAssociatedTask](taskrequestdeclineitem-getassociatedtask-method-outlook.md)|
|[GetConversation](taskrequestdeclineitem-getconversation-method-outlook.md)|
|[Move](taskrequestdeclineitem-move-method-outlook.md)|
|[PrintOut](taskrequestdeclineitem-printout-method-outlook.md)|
|[Save](taskrequestdeclineitem-save-method-outlook.md)|
|[SaveAs](taskrequestdeclineitem-saveas-method-outlook.md)|
|[ShowCategoriesDialog](taskrequestdeclineitem-showcategoriesdialog-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](taskrequestdeclineitem-actions-property-outlook.md)|
|[Application](taskrequestdeclineitem-application-property-outlook.md)|
|[Attachments](taskrequestdeclineitem-attachments-property-outlook.md)|
|[AutoResolvedWinner](taskrequestdeclineitem-autoresolvedwinner-property-outlook.md)|
|[BillingInformation](taskrequestdeclineitem-billinginformation-property-outlook.md)|
|[Body](taskrequestdeclineitem-body-property-outlook.md)|
|[Categories](taskrequestdeclineitem-categories-property-outlook.md)|
|[Class](taskrequestdeclineitem-class-property-outlook.md)|
|[Companies](taskrequestdeclineitem-companies-property-outlook.md)|
|[Conflicts](taskrequestdeclineitem-conflicts-property-outlook.md)|
|[ConversationID](taskrequestdeclineitem-conversationid-property-outlook.md)|
|[ConversationIndex](taskrequestdeclineitem-conversationindex-property-outlook.md)|
|[ConversationTopic](taskrequestdeclineitem-conversationtopic-property-outlook.md)|
|[CreationTime](taskrequestdeclineitem-creationtime-property-outlook.md)|
|[DownloadState](taskrequestdeclineitem-downloadstate-property-outlook.md)|
|[EntryID](taskrequestdeclineitem-entryid-property-outlook.md)|
|[FormDescription](taskrequestdeclineitem-formdescription-property-outlook.md)|
|[GetInspector](taskrequestdeclineitem-getinspector-property-outlook.md)|
|[Importance](taskrequestdeclineitem-importance-property-outlook.md)|
|[IsConflict](taskrequestdeclineitem-isconflict-property-outlook.md)|
|[ItemProperties](taskrequestdeclineitem-itemproperties-property-outlook.md)|
|[LastModificationTime](taskrequestdeclineitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](taskrequestdeclineitem-markfordownload-property-outlook.md)|
|[MessageClass](taskrequestdeclineitem-messageclass-property-outlook.md)|
|[Mileage](taskrequestdeclineitem-mileage-property-outlook.md)|
|[NoAging](taskrequestdeclineitem-noaging-property-outlook.md)|
|[OutlookInternalVersion](taskrequestdeclineitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](taskrequestdeclineitem-outlookversion-property-outlook.md)|
|[Parent](taskrequestdeclineitem-parent-property-outlook.md)|
|[PropertyAccessor](taskrequestdeclineitem-propertyaccessor-property-outlook.md)|
|[RTFBody](taskrequestdeclineitem-rtfbody-property-outlook.md)|
|[Saved](taskrequestdeclineitem-saved-property-outlook.md)|
|[Sensitivity](taskrequestdeclineitem-sensitivity-property-outlook.md)|
|[Session](taskrequestdeclineitem-session-property-outlook.md)|
|[Size](taskrequestdeclineitem-size-property-outlook.md)|
|[Subject](taskrequestdeclineitem-subject-property-outlook.md)|
|[UnRead](taskrequestdeclineitem-unread-property-outlook.md)|
|[UserProperties](taskrequestdeclineitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
