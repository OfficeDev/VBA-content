---
title: DocumentItem Object (Outlook)
keywords: vbaol11.chm2994
f1_keywords:
- vbaol11.chm2994
ms.prod: outlook
api_name:
- Outlook.DocumentItem
ms.assetid: 7b0a6af0-6632-3ff6-841f-5b081d0d68d8
ms.date: 06/08/2017
---


# DocumentItem Object (Outlook)

Represents any document other than a Microsoft Outlook item as an item in an Outlook folder. 


## Remarks

A  **DocumentItem** object is any document other than an Outlook item as an item in an Outlook folder. In common usage, this will be an Office document but may be any type of document or executable file.

Unlike other Outlook objects, you cannot create this object.


 **Note**  When you try to programmatically add a user-defined property to a  **DocumentItem** object, you receive the following error message: "Property is read-only." This is because the Outlook object model does not support this functionality.


## Events



|**Name**|
|:-----|
|[AfterWrite](documentitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](documentitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](documentitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](documentitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](documentitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](documentitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](documentitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](documentitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](documentitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](documentitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](documentitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](documentitem-beforedelete-event-outlook.md)|
|[BeforeRead](documentitem-beforeread-event-outlook.md)|
|[Close](documentitem-close-event-outlook.md)|
|[CustomAction](documentitem-customaction-event-outlook.md)|
|[CustomPropertyChange](documentitem-custompropertychange-event-outlook.md)|
|[Forward](documentitem-forward-event-outlook.md)|
|[Open](documentitem-open-event-outlook.md)|
|[PropertyChange](documentitem-propertychange-event-outlook.md)|
|[Read](documentitem-read-event-outlook.md)|
|[ReadComplete](documentitem-readcomplete-event-outlook.md)|
|[Reply](documentitem-reply-event-outlook.md)|
|[ReplyAll](documentitem-replyall-event-outlook.md)|
|[Send](documentitem-send-event-outlook.md)|
|[Unload](documentitem-unload-event-outlook.md)|
|[Write](documentitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Close](documentitem-close-method-outlook.md)|
|[Copy](documentitem-copy-method-outlook.md)|
|[Delete](documentitem-delete-method-outlook.md)|
|[Display](documentitem-display-method-outlook.md)|
|[Move](documentitem-move-method-outlook.md)|
|[PrintOut](documentitem-printout-method-outlook.md)|
|[Save](documentitem-save-method-outlook.md)|
|[SaveAs](documentitem-saveas-method-outlook.md)|
|[ShowCategoriesDialog](documentitem-showcategoriesdialog-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](documentitem-actions-property-outlook.md)|
|[Application](documentitem-application-property-outlook.md)|
|[Attachments](documentitem-attachments-property-outlook.md)|
|[AutoResolvedWinner](documentitem-autoresolvedwinner-property-outlook.md)|
|[BillingInformation](documentitem-billinginformation-property-outlook.md)|
|[Body](documentitem-body-property-outlook.md)|
|[Categories](documentitem-categories-property-outlook.md)|
|[Class](documentitem-class-property-outlook.md)|
|[Companies](documentitem-companies-property-outlook.md)|
|[Conflicts](documentitem-conflicts-property-outlook.md)|
|[ConversationIndex](documentitem-conversationindex-property-outlook.md)|
|[ConversationTopic](documentitem-conversationtopic-property-outlook.md)|
|[CreationTime](documentitem-creationtime-property-outlook.md)|
|[DownloadState](documentitem-downloadstate-property-outlook.md)|
|[EntryID](documentitem-entryid-property-outlook.md)|
|[FormDescription](documentitem-formdescription-property-outlook.md)|
|[GetInspector](documentitem-getinspector-property-outlook.md)|
|[Importance](documentitem-importance-property-outlook.md)|
|[IsConflict](documentitem-isconflict-property-outlook.md)|
|[ItemProperties](documentitem-itemproperties-property-outlook.md)|
|[LastModificationTime](documentitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](documentitem-markfordownload-property-outlook.md)|
|[MessageClass](documentitem-messageclass-property-outlook.md)|
|[Mileage](documentitem-mileage-property-outlook.md)|
|[NoAging](documentitem-noaging-property-outlook.md)|
|[OutlookInternalVersion](documentitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](documentitem-outlookversion-property-outlook.md)|
|[Parent](documentitem-parent-property-outlook.md)|
|[PropertyAccessor](documentitem-propertyaccessor-property-outlook.md)|
|[Saved](documentitem-saved-property-outlook.md)|
|[Sensitivity](documentitem-sensitivity-property-outlook.md)|
|[Session](documentitem-session-property-outlook.md)|
|[Size](documentitem-size-property-outlook.md)|
|[Subject](documentitem-subject-property-outlook.md)|
|[UnRead](documentitem-unread-property-outlook.md)|
|[UserProperties](documentitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
