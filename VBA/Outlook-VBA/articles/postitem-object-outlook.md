---
title: PostItem Object (Outlook)
keywords: vbaol11.chm3005
f1_keywords:
- vbaol11.chm3005
ms.prod: outlook
api_name:
- Outlook.PostItem
ms.assetid: de44065d-4e93-315a-279f-7b92f09c0465
ms.date: 06/08/2017
---


# PostItem Object (Outlook)

Represents a post in a public folder that others may browse.


## Remarks

Unlike a  **[MailItem](mailitem-object-outlook.md)** object, a **PostItem** object is not sent to a recipient. You use the **[Post](postitem-post-method-outlook.md)** method, which is analogous to the **[Send](mailitem-send-method-outlook.md)** method for the **MailItem** object, to save the **PostItem** to the target public folder instead of mailing it.

Use the  **[CreateItem](application-createitem-method-outlook.md)** or **[CreateItemFromTemplate](application-createitemfromtemplate-method-outlook.md)** method to create a **PostItem** object that represents a new post.

Use  **[Items](items-object-outlook.md)** ( _index_ ), where _index_ is the index number of a post or a value used to match the default property of a post, to return a single **PostItem** object from a public folder.


## Example

The following example returns a new post.


```
Set myItem = myOlApp.CreateItem(olPostItem)
```


## Events



|**Name**|
|:-----|
|[AfterWrite](postitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](postitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](postitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](postitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](postitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](postitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](postitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](postitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](postitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](postitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](postitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](postitem-beforedelete-event-outlook.md)|
|[BeforeRead](postitem-beforeread-event-outlook.md)|
|[Close](postitem-close-event-outlook.md)|
|[CustomAction](postitem-customaction-event-outlook.md)|
|[CustomPropertyChange](postitem-custompropertychange-event-outlook.md)|
|[Forward](postitem-forward-event-outlook.md)|
|[Open](postitem-open-event-outlook.md)|
|[PropertyChange](postitem-propertychange-event-outlook.md)|
|[Read](postitem-read-event-outlook.md)|
|[ReadComplete](postitem-readcomplete-event-outlook.md)|
|[Reply](postitem-reply-event-outlook.md)|
|[ReplyAll](postitem-replyall-event-outlook.md)|
|[Send](postitem-send-event-outlook.md)|
|[Unload](postitem-unload-event-outlook.md)|
|[Write](postitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[ClearConversationIndex](postitem-clearconversationindex-method-outlook.md)|
|[ClearTaskFlag](postitem-cleartaskflag-method-outlook.md)|
|[Close](postitem-close-method-outlook.md)|
|[Copy](postitem-copy-method-outlook.md)|
|[Delete](postitem-delete-method-outlook.md)|
|[Display](postitem-display-method-outlook.md)|
|[Forward](postitem-forward-method-outlook.md)|
|[GetConversation](postitem-getconversation-method-outlook.md)|
|[MarkAsTask](postitem-markastask-method-outlook.md)|
|[Move](postitem-move-method-outlook.md)|
|[Post](postitem-post-method-outlook.md)|
|[PrintOut](postitem-printout-method-outlook.md)|
|[Reply](postitem-reply-method-outlook.md)|
|[Save](postitem-save-method-outlook.md)|
|[SaveAs](postitem-saveas-method-outlook.md)|
|[ShowCategoriesDialog](postitem-showcategoriesdialog-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](postitem-actions-property-outlook.md)|
|[Application](postitem-application-property-outlook.md)|
|[Attachments](postitem-attachments-property-outlook.md)|
|[AutoResolvedWinner](postitem-autoresolvedwinner-property-outlook.md)|
|[BillingInformation](postitem-billinginformation-property-outlook.md)|
|[Body](postitem-body-property-outlook.md)|
|[BodyFormat](postitem-bodyformat-property-outlook.md)|
|[Categories](postitem-categories-property-outlook.md)|
|[Class](postitem-class-property-outlook.md)|
|[Companies](postitem-companies-property-outlook.md)|
|[Conflicts](postitem-conflicts-property-outlook.md)|
|[ConversationID](postitem-conversationid-property-outlook.md)|
|[ConversationIndex](postitem-conversationindex-property-outlook.md)|
|[ConversationTopic](postitem-conversationtopic-property-outlook.md)|
|[CreationTime](postitem-creationtime-property-outlook.md)|
|[DownloadState](postitem-downloadstate-property-outlook.md)|
|[EntryID](postitem-entryid-property-outlook.md)|
|[ExpiryTime](postitem-expirytime-property-outlook.md)|
|[FormDescription](postitem-formdescription-property-outlook.md)|
|[GetInspector](postitem-getinspector-property-outlook.md)|
|[HTMLBody](postitem-htmlbody-property-outlook.md)|
|[Importance](postitem-importance-property-outlook.md)|
|[InternetCodepage](postitem-internetcodepage-property-outlook.md)|
|[IsConflict](postitem-isconflict-property-outlook.md)|
|[IsMarkedAsTask](postitem-ismarkedastask-property-outlook.md)|
|[ItemProperties](postitem-itemproperties-property-outlook.md)|
|[LastModificationTime](postitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](postitem-markfordownload-property-outlook.md)|
|[MessageClass](postitem-messageclass-property-outlook.md)|
|[Mileage](postitem-mileage-property-outlook.md)|
|[NoAging](postitem-noaging-property-outlook.md)|
|[OutlookInternalVersion](postitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](postitem-outlookversion-property-outlook.md)|
|[Parent](postitem-parent-property-outlook.md)|
|[PropertyAccessor](postitem-propertyaccessor-property-outlook.md)|
|[ReceivedTime](postitem-receivedtime-property-outlook.md)|
|[ReminderOverrideDefault](postitem-reminderoverridedefault-property-outlook.md)|
|[ReminderPlaySound](postitem-reminderplaysound-property-outlook.md)|
|[ReminderSet](postitem-reminderset-property-outlook.md)|
|[ReminderSoundFile](postitem-remindersoundfile-property-outlook.md)|
|[ReminderTime](postitem-remindertime-property-outlook.md)|
|[RTFBody](postitem-rtfbody-property-outlook.md)|
|[Saved](postitem-saved-property-outlook.md)|
|[SenderEmailAddress](postitem-senderemailaddress-property-outlook.md)|
|[SenderEmailType](postitem-senderemailtype-property-outlook.md)|
|[SenderName](postitem-sendername-property-outlook.md)|
|[Sensitivity](postitem-sensitivity-property-outlook.md)|
|[SentOn](postitem-senton-property-outlook.md)|
|[Session](postitem-session-property-outlook.md)|
|[Size](postitem-size-property-outlook.md)|
|[Subject](postitem-subject-property-outlook.md)|
|[TaskCompletedDate](postitem-taskcompleteddate-property-outlook.md)|
|[TaskDueDate](postitem-taskduedate-property-outlook.md)|
|[TaskStartDate](postitem-taskstartdate-property-outlook.md)|
|[TaskSubject](postitem-tasksubject-property-outlook.md)|
|[ToDoTaskOrdinal](postitem-todotaskordinal-property-outlook.md)|
|[UnRead](postitem-unread-property-outlook.md)|
|[UserProperties](postitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
