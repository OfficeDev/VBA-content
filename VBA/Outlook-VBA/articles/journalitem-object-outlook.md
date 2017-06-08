---
title: JournalItem Object (Outlook)
keywords: vbaol11.chm2999
f1_keywords:
- vbaol11.chm2999
ms.prod: outlook
api_name:
- Outlook.JournalItem
ms.assetid: 6e850295-39f9-47b8-e866-9622e9958c69
ms.date: 06/08/2017
---


# JournalItem Object (Outlook)

Represents a journal entry in a Journal folder. 


## Remarks

A journal entry represents a record of all Outlook-moderated transactions for any given period.

Use the  **[CreateItem](application-createitem-method-outlook.md)** method to create a **JournalItem** object that represents a new journal entry.

Use  **[Items](folder-items-property-outlook.md)** ( _index_ ), where _index_ is the index number of a journal entry or a value used to match the default property of a journal entry, to return a single **JournalItem** object from a Journal folder.


## Example

The following Visual Basic for Applications (VBA) example returns a new journal entry.


```
Set myItem = Application.CreateItem(olJournalItem)
```


## Events



|**Name**|
|:-----|
|[AfterWrite](journalitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](journalitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](journalitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](journalitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](journalitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](journalitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](journalitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](journalitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](journalitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](journalitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](journalitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](journalitem-beforedelete-event-outlook.md)|
|[BeforeRead](journalitem-beforeread-event-outlook.md)|
|[Close](journalitem-close-event-outlook.md)|
|[CustomAction](journalitem-customaction-event-outlook.md)|
|[CustomPropertyChange](journalitem-custompropertychange-event-outlook.md)|
|[Forward](journalitem-forward-event-outlook.md)|
|[Open](journalitem-open-event-outlook.md)|
|[PropertyChange](journalitem-propertychange-event-outlook.md)|
|[Read](journalitem-read-event-outlook.md)|
|[ReadComplete](journalitem-readcomplete-event-outlook.md)|
|[Reply](journalitem-reply-event-outlook.md)|
|[ReplyAll](journalitem-replyall-event-outlook.md)|
|[Send](journalitem-send-event-outlook.md)|
|[Unload](journalitem-unload-event-outlook.md)|
|[Write](journalitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Close](journalitem-close-method-outlook.md)|
|[Copy](journalitem-copy-method-outlook.md)|
|[Delete](journalitem-delete-method-outlook.md)|
|[Display](journalitem-display-method-outlook.md)|
|[Forward](journalitem-forward-method-outlook.md)|
|[GetConversation](journalitem-getconversation-method-outlook.md)|
|[Move](journalitem-move-method-outlook.md)|
|[PrintOut](journalitem-printout-method-outlook.md)|
|[Reply](journalitem-reply-method-outlook.md)|
|[ReplyAll](journalitem-replyall-method-outlook.md)|
|[Save](journalitem-save-method-outlook.md)|
|[SaveAs](journalitem-saveas-method-outlook.md)|
|[ShowCategoriesDialog](journalitem-showcategoriesdialog-method-outlook.md)|
|[StartTimer](journalitem-starttimer-method-outlook.md)|
|[StopTimer](journalitem-stoptimer-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](journalitem-actions-property-outlook.md)|
|[Application](journalitem-application-property-outlook.md)|
|[Attachments](journalitem-attachments-property-outlook.md)|
|[AutoResolvedWinner](journalitem-autoresolvedwinner-property-outlook.md)|
|[BillingInformation](journalitem-billinginformation-property-outlook.md)|
|[Body](journalitem-body-property-outlook.md)|
|[Categories](journalitem-categories-property-outlook.md)|
|[Class](journalitem-class-property-outlook.md)|
|[Companies](journalitem-companies-property-outlook.md)|
|[Conflicts](journalitem-conflicts-property-outlook.md)|
|[ContactNames](journalitem-contactnames-property-outlook.md)|
|[ConversationID](journalitem-conversationid-property-outlook.md)|
|[ConversationIndex](journalitem-conversationindex-property-outlook.md)|
|[ConversationTopic](journalitem-conversationtopic-property-outlook.md)|
|[CreationTime](journalitem-creationtime-property-outlook.md)|
|[DocPosted](journalitem-docposted-property-outlook.md)|
|[DocPrinted](journalitem-docprinted-property-outlook.md)|
|[DocRouted](journalitem-docrouted-property-outlook.md)|
|[DocSaved](journalitem-docsaved-property-outlook.md)|
|[DownloadState](journalitem-downloadstate-property-outlook.md)|
|[Duration](journalitem-duration-property-outlook.md)|
|[End](journalitem-end-property-outlook.md)|
|[EntryID](journalitem-entryid-property-outlook.md)|
|[FormDescription](journalitem-formdescription-property-outlook.md)|
|[GetInspector](journalitem-getinspector-property-outlook.md)|
|[Importance](journalitem-importance-property-outlook.md)|
|[IsConflict](journalitem-isconflict-property-outlook.md)|
|[ItemProperties](journalitem-itemproperties-property-outlook.md)|
|[LastModificationTime](journalitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](journalitem-markfordownload-property-outlook.md)|
|[MessageClass](journalitem-messageclass-property-outlook.md)|
|[Mileage](journalitem-mileage-property-outlook.md)|
|[NoAging](journalitem-noaging-property-outlook.md)|
|[OutlookInternalVersion](journalitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](journalitem-outlookversion-property-outlook.md)|
|[Parent](journalitem-parent-property-outlook.md)|
|[PropertyAccessor](journalitem-propertyaccessor-property-outlook.md)|
|[Recipients](journalitem-recipients-property-outlook.md)|
|[Saved](journalitem-saved-property-outlook.md)|
|[Sensitivity](journalitem-sensitivity-property-outlook.md)|
|[Session](journalitem-session-property-outlook.md)|
|[Size](journalitem-size-property-outlook.md)|
|[Start](journalitem-start-property-outlook.md)|
|[Subject](journalitem-subject-property-outlook.md)|
|[Type](journalitem-type-property-outlook.md)|
|[UnRead](journalitem-unread-property-outlook.md)|
|[UserProperties](journalitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
