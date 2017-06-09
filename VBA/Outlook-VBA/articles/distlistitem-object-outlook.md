---
title: DistListItem Object (Outlook)
keywords: vbaol11.chm2993
f1_keywords:
- vbaol11.chm2993
ms.prod: outlook
api_name:
- Outlook.DistListItem
ms.assetid: 027c3986-abff-d9b1-ecc2-26d60805e952
ms.date: 06/08/2017
---


# DistListItem Object (Outlook)

Represents a distribution list in a Contacts folder.


## Remarks

 A distribution list can contain multiple recipients and is used to send messages to everyone in the list.

Use the  **[CreateItem](application-createitem-method-outlook.md)** method to create a **DistListItem** object that represents a new distribution list.

Use  **[Items](folder-items-property-outlook.md)** ( _index_ ), where _index_ is the index number of an item in a contacts folder or a value used to match the default property of an item in the folder, to return a single **DistListItem** object from a contacts folder (that is, a folder whose default item type is **olContactItem** ).


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates and displays a new distribution list.


```
Set myItem = Application.CreateItem(olDistributionListItem) 
 
myItem.Display
```

The following Visual Basic for Applications example sets the current folder as the contacts folder and displays an existing distribution list named Project Team in the folder.




```
Set myNamespace = Application.GetNamespace("MAPI") 
 
Set myFolder = myNamespace.GetDefaultFolder(olFolderContacts) 
 
myFolder.Display 
 
Set myItem = myFolder.Items("Project Team") 
 
myItem.Display
```


## Events



|**Name**|
|:-----|
|[AfterWrite](distlistitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](distlistitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](distlistitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](distlistitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](distlistitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](distlistitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](distlistitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](distlistitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](distlistitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](distlistitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](distlistitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](distlistitem-beforedelete-event-outlook.md)|
|[BeforeRead](distlistitem-beforeread-event-outlook.md)|
|[Close](distlistitem-close-event-outlook.md)|
|[CustomAction](distlistitem-customaction-event-outlook.md)|
|[CustomPropertyChange](distlistitem-custompropertychange-event-outlook.md)|
|[Forward](distlistitem-forward-event-outlook.md)|
|[Open](distlistitem-open-event-outlook.md)|
|[PropertyChange](distlistitem-propertychange-event-outlook.md)|
|[Read](distlistitem-read-event-outlook.md)|
|[ReadComplete](distlistitem-readcomplete-event-outlook.md)|
|[Reply](distlistitem-reply-event-outlook.md)|
|[ReplyAll](distlistitem-replyall-event-outlook.md)|
|[Send](distlistitem-send-event-outlook.md)|
|[Unload](distlistitem-unload-event-outlook.md)|
|[Write](distlistitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[AddMember](distlistitem-addmember-method-outlook.md)|
|[AddMembers](distlistitem-addmembers-method-outlook.md)|
|[ClearTaskFlag](distlistitem-cleartaskflag-method-outlook.md)|
|[Close](distlistitem-close-method-outlook.md)|
|[Copy](distlistitem-copy-method-outlook.md)|
|[Delete](distlistitem-delete-method-outlook.md)|
|[Display](distlistitem-display-method-outlook.md)|
|[GetConversation](distlistitem-getconversation-method-outlook.md)|
|[GetMember](distlistitem-getmember-method-outlook.md)|
|[MarkAsTask](distlistitem-markastask-method-outlook.md)|
|[Move](distlistitem-move-method-outlook.md)|
|[PrintOut](distlistitem-printout-method-outlook.md)|
|[RemoveMember](distlistitem-removemember-method-outlook.md)|
|[RemoveMembers](distlistitem-removemembers-method-outlook.md)|
|[Save](distlistitem-save-method-outlook.md)|
|[SaveAs](distlistitem-saveas-method-outlook.md)|
|[ShowCategoriesDialog](distlistitem-showcategoriesdialog-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](distlistitem-actions-property-outlook.md)|
|[Application](distlistitem-application-property-outlook.md)|
|[Attachments](distlistitem-attachments-property-outlook.md)|
|[AutoResolvedWinner](distlistitem-autoresolvedwinner-property-outlook.md)|
|[BillingInformation](distlistitem-billinginformation-property-outlook.md)|
|[Body](distlistitem-body-property-outlook.md)|
|[Categories](distlistitem-categories-property-outlook.md)|
|[Class](distlistitem-class-property-outlook.md)|
|[Companies](distlistitem-companies-property-outlook.md)|
|[Conflicts](distlistitem-conflicts-property-outlook.md)|
|[ConversationID](distlistitem-conversationid-property-outlook.md)|
|[ConversationIndex](distlistitem-conversationindex-property-outlook.md)|
|[ConversationTopic](distlistitem-conversationtopic-property-outlook.md)|
|[CreationTime](distlistitem-creationtime-property-outlook.md)|
|[DLName](distlistitem-dlname-property-outlook.md)|
|[DownloadState](distlistitem-downloadstate-property-outlook.md)|
|[EntryID](distlistitem-entryid-property-outlook.md)|
|[FormDescription](distlistitem-formdescription-property-outlook.md)|
|[GetInspector](distlistitem-getinspector-property-outlook.md)|
|[Importance](distlistitem-importance-property-outlook.md)|
|[IsConflict](distlistitem-isconflict-property-outlook.md)|
|[IsMarkedAsTask](distlistitem-ismarkedastask-property-outlook.md)|
|[ItemProperties](distlistitem-itemproperties-property-outlook.md)|
|[LastModificationTime](distlistitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](distlistitem-markfordownload-property-outlook.md)|
|[MemberCount](distlistitem-membercount-property-outlook.md)|
|[MessageClass](distlistitem-messageclass-property-outlook.md)|
|[Mileage](distlistitem-mileage-property-outlook.md)|
|[NoAging](distlistitem-noaging-property-outlook.md)|
|[OutlookInternalVersion](distlistitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](distlistitem-outlookversion-property-outlook.md)|
|[Parent](distlistitem-parent-property-outlook.md)|
|[PropertyAccessor](distlistitem-propertyaccessor-property-outlook.md)|
|[ReminderOverrideDefault](distlistitem-reminderoverridedefault-property-outlook.md)|
|[ReminderPlaySound](distlistitem-reminderplaysound-property-outlook.md)|
|[ReminderSet](distlistitem-reminderset-property-outlook.md)|
|[ReminderSoundFile](distlistitem-remindersoundfile-property-outlook.md)|
|[ReminderTime](distlistitem-remindertime-property-outlook.md)|
|[RTFBody](distlistitem-rtfbody-property-outlook.md)|
|[Saved](distlistitem-saved-property-outlook.md)|
|[Sensitivity](distlistitem-sensitivity-property-outlook.md)|
|[Session](distlistitem-session-property-outlook.md)|
|[Size](distlistitem-size-property-outlook.md)|
|[Subject](distlistitem-subject-property-outlook.md)|
|[TaskCompletedDate](distlistitem-taskcompleteddate-property-outlook.md)|
|[TaskDueDate](distlistitem-taskduedate-property-outlook.md)|
|[TaskStartDate](distlistitem-taskstartdate-property-outlook.md)|
|[TaskSubject](distlistitem-tasksubject-property-outlook.md)|
|[ToDoTaskOrdinal](distlistitem-todotaskordinal-property-outlook.md)|
|[UnRead](distlistitem-unread-property-outlook.md)|
|[UserProperties](distlistitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
