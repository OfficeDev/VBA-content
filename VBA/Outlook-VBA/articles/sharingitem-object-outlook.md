---
title: SharingItem Object (Outlook)
keywords: vbaol11.chm3016
f1_keywords:
- vbaol11.chm3016
ms.prod: outlook
api_name:
- Outlook.SharingItem
ms.assetid: 63dd3451-44f3-7cc4-c6e2-7dad5835a7d2
ms.date: 06/08/2017
---


# SharingItem Object (Outlook)

Represents a sharing message in an Inbox folder.


## Remarks

Use the  **[CreateSharingItem](namespace-createsharingitem-method-outlook.md)** method of the **[NameSpace](namespace-object-outlook.md)** object to create a **SharingItem** object that represents a new sharing request or sharing invitation.

Use  **[Item](folders-item-method-outlook.md)** ( _index_ ), where _index_ is the index number of a sharing message or a value used to match the default property of a message, to return a single **SharingItem** object from an Inbox folder.


## Example

The following Visual Basic for Applications (VBA) example creates and displays a new sharing invitation for the Tasks folder.


```
Public Sub CreateTasksSharingItem() 
 
 
 
 Dim oNamespace As NameSpace 
 
 Dim oFolder As Folder 
 
 Dim oSharingItem As SharingItem 
 
 
 
 On Error GoTo ErrRoutine 
 
 
 
 Set oNamespace = Application.GetNamespace("MAPI") 
 
 Set oFolder = oNamespace.GetDefaultFolder(olFolderTasks) 
 
 Set oSharingItem = oNamespace.CreateSharingItem(oFolder) 
 
 
 
 oSharingItem.Display 
 
 
 
EndRoutine: 
 
 On Error GoTo 0 
 
 Set oSharingItem = Nothing 
 
 Set oFolder = Nothing 
 
 Set oNamespace = Nothing 
 
Exit Sub 
 
 
 
ErrRoutine: 
 
 MsgBox Err.Description, _ 
 
 vbOKOnly, _ 
 
 Err.Number &amp; " - " &amp; Err.Source 
 
 GoTo EndRoutine 
 
End Sub 
 

```


## Events



|**Name**|
|:-----|
|[AfterWrite](sharingitem-afterwrite-event-outlook.md)|
|[AttachmentAdd](sharingitem-attachmentadd-event-outlook.md)|
|[AttachmentRead](sharingitem-attachmentread-event-outlook.md)|
|[AttachmentRemove](sharingitem-attachmentremove-event-outlook.md)|
|[BeforeAttachmentAdd](sharingitem-beforeattachmentadd-event-outlook.md)|
|[BeforeAttachmentPreview](sharingitem-beforeattachmentpreview-event-outlook.md)|
|[BeforeAttachmentRead](sharingitem-beforeattachmentread-event-outlook.md)|
|[BeforeAttachmentSave](sharingitem-beforeattachmentsave-event-outlook.md)|
|[BeforeAttachmentWriteToTempFile](sharingitem-beforeattachmentwritetotempfile-event-outlook.md)|
|[BeforeAutoSave](sharingitem-beforeautosave-event-outlook.md)|
|[BeforeCheckNames](sharingitem-beforechecknames-event-outlook.md)|
|[BeforeDelete](sharingitem-beforedelete-event-outlook.md)|
|[BeforeRead](sharingitem-beforeread-event-outlook.md)|
|[Close](sharingitem-close-event-outlook.md)|
|[CustomAction](sharingitem-customaction-event-outlook.md)|
|[CustomPropertyChange](sharingitem-custompropertychange-event-outlook.md)|
|[Forward](sharingitem-forward-event-outlook.md)|
|[Open](sharingitem-open-event-outlook.md)|
|[PropertyChange](sharingitem-propertychange-event-outlook.md)|
|[Read](sharingitem-read-event-outlook.md)|
|[ReadComplete](sharingitem-readcomplete-event-outlook.md)|
|[Reply](sharingitem-reply-event-outlook.md)|
|[ReplyAll](sharingitem-replyall-event-outlook.md)|
|[Send](sharingitem-send-event-outlook.md)|
|[Unload](sharingitem-unload-event-outlook.md)|
|[Write](sharingitem-write-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[AddBusinessCard](sharingitem-addbusinesscard-method-outlook.md)|
|[Allow](sharingitem-allow-method-outlook.md)|
|[ClearConversationIndex](sharingitem-clearconversationindex-method-outlook.md)|
|[ClearTaskFlag](sharingitem-cleartaskflag-method-outlook.md)|
|[Close](sharingitem-close-method-outlook.md)|
|[Copy](sharingitem-copy-method-outlook.md)|
|[Delete](sharingitem-delete-method-outlook.md)|
|[Deny](sharingitem-deny-method-outlook.md)|
|[Display](sharingitem-display-method-outlook.md)|
|[Forward](sharingitem-forward-method-outlook.md)|
|[GetConversation](sharingitem-getconversation-method-outlook.md)|
|[MarkAsTask](sharingitem-markastask-method-outlook.md)|
|[Move](sharingitem-move-method-outlook.md)|
|[OpenSharedFolder](sharingitem-opensharedfolder-method-outlook.md)|
|[PrintOut](sharingitem-printout-method-outlook.md)|
|[Reply](sharingitem-reply-method-outlook.md)|
|[ReplyAll](sharingitem-replyall-method-outlook.md)|
|[Save](sharingitem-save-method-outlook.md)|
|[SaveAs](sharingitem-saveas-method-outlook.md)|
|[Send](sharingitem-send-method-outlook.md)|
|[ShowCategoriesDialog](sharingitem-showcategoriesdialog-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Actions](sharingitem-actions-property-outlook.md)|
|[AllowWriteAccess](sharingitem-allowwriteaccess-property-outlook.md)|
|[AlternateRecipientAllowed](sharingitem-alternaterecipientallowed-property-outlook.md)|
|[Application](sharingitem-application-property-outlook.md)|
|[Attachments](sharingitem-attachments-property-outlook.md)|
|[AutoForwarded](sharingitem-autoforwarded-property-outlook.md)|
|[BCC](sharingitem-bcc-property-outlook.md)|
|[BillingInformation](sharingitem-billinginformation-property-outlook.md)|
|[Body](sharingitem-body-property-outlook.md)|
|[BodyFormat](sharingitem-bodyformat-property-outlook.md)|
|[Categories](sharingitem-categories-property-outlook.md)|
|[CC](sharingitem-cc-property-outlook.md)|
|[Class](sharingitem-class-property-outlook.md)|
|[Companies](sharingitem-companies-property-outlook.md)|
|[Conflicts](sharingitem-conflicts-property-outlook.md)|
|[ConversationID](sharingitem-conversationid-property-outlook.md)|
|[ConversationIndex](sharingitem-conversationindex-property-outlook.md)|
|[ConversationTopic](sharingitem-conversationtopic-property-outlook.md)|
|[CreationTime](sharingitem-creationtime-property-outlook.md)|
|[DeferredDeliveryTime](sharingitem-deferreddeliverytime-property-outlook.md)|
|[DeleteAfterSubmit](sharingitem-deleteaftersubmit-property-outlook.md)|
|[DownloadState](sharingitem-downloadstate-property-outlook.md)|
|[EntryID](sharingitem-entryid-property-outlook.md)|
|[ExpiryTime](sharingitem-expirytime-property-outlook.md)|
|[FlagRequest](sharingitem-flagrequest-property-outlook.md)|
|[FormDescription](sharingitem-formdescription-property-outlook.md)|
|[GetInspector](sharingitem-getinspector-property-outlook.md)|
|[HTMLBody](sharingitem-htmlbody-property-outlook.md)|
|[Importance](sharingitem-importance-property-outlook.md)|
|[InternetCodepage](sharingitem-internetcodepage-property-outlook.md)|
|[IsConflict](sharingitem-isconflict-property-outlook.md)|
|[IsMarkedAsTask](sharingitem-ismarkedastask-property-outlook.md)|
|[ItemProperties](sharingitem-itemproperties-property-outlook.md)|
|[LastModificationTime](sharingitem-lastmodificationtime-property-outlook.md)|
|[MarkForDownload](sharingitem-markfordownload-property-outlook.md)|
|[MessageClass](sharingitem-messageclass-property-outlook.md)|
|[Mileage](sharingitem-mileage-property-outlook.md)|
|[NoAging](sharingitem-noaging-property-outlook.md)|
|[OriginatorDeliveryReportRequested](sharingitem-originatordeliveryreportrequested-property-outlook.md)|
|[OutlookInternalVersion](sharingitem-outlookinternalversion-property-outlook.md)|
|[OutlookVersion](sharingitem-outlookversion-property-outlook.md)|
|[Parent](sharingitem-parent-property-outlook.md)|
|[Permission](sharingitem-permission-property-outlook.md)|
|[PermissionService](sharingitem-permissionservice-property-outlook.md)|
|[PermissionTemplateGuid](sharingitem-permissiontemplateguid-property-outlook.md)|
|[PropertyAccessor](sharingitem-propertyaccessor-property-outlook.md)|
|[ReadReceiptRequested](sharingitem-readreceiptrequested-property-outlook.md)|
|[ReceivedByEntryID](sharingitem-receivedbyentryid-property-outlook.md)|
|[ReceivedByName](sharingitem-receivedbyname-property-outlook.md)|
|[ReceivedOnBehalfOfEntryID](sharingitem-receivedonbehalfofentryid-property-outlook.md)|
|[ReceivedOnBehalfOfName](sharingitem-receivedonbehalfofname-property-outlook.md)|
|[ReceivedTime](sharingitem-receivedtime-property-outlook.md)|
|[RecipientReassignmentProhibited](sharingitem-recipientreassignmentprohibited-property-outlook.md)|
|[Recipients](sharingitem-recipients-property-outlook.md)|
|[ReminderOverrideDefault](sharingitem-reminderoverridedefault-property-outlook.md)|
|[ReminderPlaySound](sharingitem-reminderplaysound-property-outlook.md)|
|[ReminderSet](sharingitem-reminderset-property-outlook.md)|
|[ReminderSoundFile](sharingitem-remindersoundfile-property-outlook.md)|
|[ReminderTime](sharingitem-remindertime-property-outlook.md)|
|[RemoteID](sharingitem-remoteid-property-outlook.md)|
|[RemoteName](sharingitem-remotename-property-outlook.md)|
|[RemotePath](sharingitem-remotepath-property-outlook.md)|
|[RemoteStatus](sharingitem-remotestatus-property-outlook.md)|
|[ReplyRecipientNames](sharingitem-replyrecipientnames-property-outlook.md)|
|[ReplyRecipients](sharingitem-replyrecipients-property-outlook.md)|
|[RequestedFolder](sharingitem-requestedfolder-property-outlook.md)|
|[RetentionExpirationDate](sharingitem-retentionexpirationdate-property-outlook.md)|
|[RetentionPolicyName](sharingitem-retentionpolicyname-property-outlook.md)|
|[RTFBody](sharingitem-rtfbody-property-outlook.md)|
|[Saved](sharingitem-saved-property-outlook.md)|
|[SaveSentMessageFolder](sharingitem-savesentmessagefolder-property-outlook.md)|
|[SenderEmailAddress](sharingitem-senderemailaddress-property-outlook.md)|
|[SenderEmailType](sharingitem-senderemailtype-property-outlook.md)|
|[SenderName](sharingitem-sendername-property-outlook.md)|
|[SendUsingAccount](sharingitem-sendusingaccount-property-outlook.md)|
|[Sensitivity](sharingitem-sensitivity-property-outlook.md)|
|[Sent](sharingitem-sent-property-outlook.md)|
|[SentOn](sharingitem-senton-property-outlook.md)|
|[SentOnBehalfOfName](sharingitem-sentonbehalfofname-property-outlook.md)|
|[Session](sharingitem-session-property-outlook.md)|
|[SharingProvider](sharingitem-sharingprovider-property-outlook.md)|
|[SharingProviderGuid](sharingitem-sharingproviderguid-property-outlook.md)|
|[Size](sharingitem-size-property-outlook.md)|
|[Subject](sharingitem-subject-property-outlook.md)|
|[Submitted](sharingitem-submitted-property-outlook.md)|
|[TaskCompletedDate](sharingitem-taskcompleteddate-property-outlook.md)|
|[TaskDueDate](sharingitem-taskduedate-property-outlook.md)|
|[TaskStartDate](sharingitem-taskstartdate-property-outlook.md)|
|[TaskSubject](sharingitem-tasksubject-property-outlook.md)|
|[To](sharingitem-to-property-outlook.md)|
|[ToDoTaskOrdinal](sharingitem-todotaskordinal-property-outlook.md)|
|[Type](sharingitem-type-property-outlook.md)|
|[UnRead](sharingitem-unread-property-outlook.md)|
|[UserProperties](sharingitem-userproperties-property-outlook.md)|

## See also


#### Other resources


[SharingItem Object Members](http://msdn.microsoft.com/library/719ad60e-2242-2c54-778f-006b61690389%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
