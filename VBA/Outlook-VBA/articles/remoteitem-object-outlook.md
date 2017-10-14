---
title: RemoteItem Object (Outlook)
keywords: vbaol11.chm3006
f1_keywords:
- vbaol11.chm3006
ms.prod: outlook
api_name:
- Outlook.RemoteItem
ms.assetid: 6302aaff-cdcf-4d86-60f1-4bed15540d9f
ms.date: 06/08/2017
---


# RemoteItem Object (Outlook)

Represents a remote item in an Inbox folder.


## Remarks

The  **RemoteItem** object is similar to the **[MailItem](mailitem-object-outlook.md)** object, but it contains only the **Subject**,  **Received Date** and **Time**,  **Sender**,  **Size**, and the first 256 characters of the body of the message. It is used to give someone connecting in remote mode enough information to decide whether or not to download the corresponding mail message. However, the headers in items contained in an Offline Folders file (.ost) cannot be accessed using the  **RemoteItem** object.

Unlike other Microsoft Outlook objects, you cannot create this object. Remote items are created by Outlook automatically when you use a Remote Access System (RAS) connection. Each  **RemoteItem** object created on the local system corresponds to a preexisting **MailItem** object on the remote system.

The  **RemoteItem** object inherits a number of properties, methods, and events that, because of the nature of the object, have no function. The **Object Browser** shows these properties, methods, and events as belonging to the **RemoteItem** object, but trying to use them will produce no effect.

The methods that do not work for the  **RemoteItem** object include **Close**, **Copy**, **Display**, **Move**, and **Save**.

The properties that do not work for the  **RemoteItem** object include **BillingInformation**, **Body**, **Categories**, **Companies**, and **Mileage**.

The events that do not work for the  **RemoteItem** object include **Open**, **Close**, **Forward**, **Reply**, **ReplyAll**, and **Send**.


## Events



|**Name**|
|:-----|
|[AfterWrite](remoteitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](remoteitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](remoteitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](remoteitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](remoteitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](remoteitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](remoteitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](remoteitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](remoteitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](remoteitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](remoteitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](remoteitem-beforedelete-event-outlook.md)|
|[BeforeRead](remoteitem-beforeread-event-outlook.md)|
|[Close](remoteitem-close-event-outlook.md)|
|[CustomAction](remoteitem-customaction-event-outlook.md)|
|[CustomPropertyChange](remoteitem-custompropertychange-event-outlook.md)|
|[Forward](remoteitem-forward-event-outlook.md)|
|[Open](remoteitem-open-event-outlook.md)|
|[PropertyChange](remoteitem-propertychange-event-outlook.md)|
|[Read](remoteitem-read-event-outlook.md)|
|[ReadComplete](remoteitem-readcomplete-event-outlook.md)|
|[Reply](remoteitem-reply-event-outlook.md)|
|[ReplyAll](remoteitem-replyall-event-outlook.md)|
|[Send](remoteitem-send-event-outlook.md)|
|[Unload](remoteitem-unload-event-outlook.md)|
|[Write](remoteitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Close](remoteitem-close-method-outlook.md)|
|[Copy](remoteitem-copy-method-outlook.md)|
|[Delete](remoteitem-delete-method-outlook.md)|
|[Display](remoteitem-display-method-outlook.md)|
|[GetConversation](remoteitem-getconversation-method-outlook.md)|
|[Move](remoteitem-move-method-outlook.md)|
|[PrintOut](remoteitem-printout-method-outlook.md)|
|[Save](remoteitem-save-method-outlook.md)|
|[SaveAs](remoteitem-saveas-method-outlook.md)|
|[ShowCategoriesDialog](remoteitem-showcategoriesdialog-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](remoteitem-actions-property-outlook.md)|
|[Application](remoteitem-application-property-outlook.md)|
|[Attachments](remoteitem-attachments-property-outlook.md)|
|[AutoResolvedWinner](remoteitem-autoresolvedwinner-property-outlook.md)|
|[BillingInformation](remoteitem-billinginformation-property-outlook.md)|
|[Body](remoteitem-body-property-outlook.md)|
|[Categories](remoteitem-categories-property-outlook.md)|
|[Class](remoteitem-class-property-outlook.md)|
|[Companies](remoteitem-companies-property-outlook.md)|
|[Conflicts](remoteitem-conflicts-property-outlook.md)|
|[ConversationID](remoteitem-conversationid-property-outlook.md)|
|[ConversationIndex](remoteitem-conversationindex-property-outlook.md)|
|[ConversationTopic](remoteitem-conversationtopic-property-outlook.md)|
|[CreationTime](remoteitem-creationtime-property-outlook.md)|
|[DownloadState](remoteitem-downloadstate-property-outlook.md)|
|[EntryID](remoteitem-entryid-property-outlook.md)|
|[FormDescription](remoteitem-formdescription-property-outlook.md)|
|[GetInspector](remoteitem-getinspector-property-outlook.md)|
|[HasAttachment](remoteitem-hasattachment-property-outlook.md)|
|[Importance](remoteitem-importance-property-outlook.md)|
|[IsConflict](remoteitem-isconflict-property-outlook.md)|
|[ItemProperties](remoteitem-itemproperties-property-outlook.md)|
|[LastModificationTime](remoteitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](remoteitem-markfordownload-property-outlook.md)|
|[MessageClass](remoteitem-messageclass-property-outlook.md)|
|[Mileage](remoteitem-mileage-property-outlook.md)|
|[NoAging](remoteitem-noaging-property-outlook.md)|
|[OutlookInternalVersion](remoteitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](remoteitem-outlookversion-property-outlook.md)|
|[Parent](remoteitem-parent-property-outlook.md)|
|[PropertyAccessor](remoteitem-propertyaccessor-property-outlook.md)|
|[RemoteMessageClass](remoteitem-remotemessageclass-property-outlook.md)|
|[Saved](remoteitem-saved-property-outlook.md)|
|[Sensitivity](remoteitem-sensitivity-property-outlook.md)|
|[Session](remoteitem-session-property-outlook.md)|
|[Size](remoteitem-size-property-outlook.md)|
|[Subject](remoteitem-subject-property-outlook.md)|
|[TransferSize](remoteitem-transfersize-property-outlook.md)|
|[TransferTime](remoteitem-transfertime-property-outlook.md)|
|[UnRead](remoteitem-unread-property-outlook.md)|
|[UserProperties](remoteitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
