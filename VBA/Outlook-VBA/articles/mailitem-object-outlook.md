---
title: MailItem Object (Outlook)
keywords: vbaol11.chm2987
f1_keywords:
- vbaol11.chm2987
ms.prod: outlook
api_name:
- Outlook.MailItem
ms.assetid: 14197346-05d2-0250-fa4c-4a6b07daf25f
ms.date: 06/08/2017
---


# MailItem Object (Outlook)

Represents a mail message.


## Remarks

Use the  **[CreateItem](http://msdn.microsoft.com/library/e5fbf367-db16-5042-823e-68e6b805e612%28Office.15%29.aspx)** method to create a **MailItem** object that represents a new mail message.

Use the  **[Folder.Items](http://msdn.microsoft.com/library/441820e7-5fe8-e5ef-83c0-9c87fd3dc9e3%28Office.15%29.aspx)** property to obtain an **[Items](http://msdn.microsoft.com/library/3a99730b-e62a-5ca6-f6ec-911c95173242%28Office.15%29.aspx)** collection representing the mail items in a folder, and the **[Items.Item](http://msdn.microsoft.com/library/89a031e0-c0a3-fc22-f485-189df8db45f4%28Office.15%29.aspx)** (_index_) method, where _index_ is the index number of a mail message or a value used to match the default property of a message, to return a single **MailItem** object from the specified folder.


## Example

The following Visual Basic for Applications (VBA) example creates and displays a new mail message.


```
Sub CreateMail() 
 
 Dim myItem As Object 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 myItem.Subject = "Mail to myself" 
 
 myItem.Display 
 
End Sub
```

The following VBA example sets the current folder as the Inbox and displays the second mail message in the folder. In general, the order of mail messages in a folder is not guaranteed to be in a particular order. 




```
Sub DisplayMail() 
 
 Dim myItem As Object 
 
 Dim myFolder As Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNamespace.GetDefaultFolder(olFolderInbox) 
 
 myFolder.Display 
 
 Set myItem = myFolder.Items(2) 
 
 myItem.Display 
 
End Sub
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/e8face1d-06bd-2799-5afd-53048bb03acd%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/ae95c10b-f8dc-0341-4153-c7805d973df9%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/9da23894-0867-aac8-2275-251e32ad4180%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/3c7fb9c8-55ef-f298-ab00-95e7537c3f1a%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/d053d72c-07fa-275e-6e1a-8d54e23119ec%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/279e1af4-38e1-d6b5-50a5-9ebd517826ae%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/00d35fff-b1d2-0da2-7315-a9fce2f28e80%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/b36eb8dc-3128-c75c-9c2d-b5321d93680c%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/fad940fa-3ab8-ac9c-0cc1-adc36c695af8%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/0c725b91-f72f-7ceb-b2a9-da4f0369cf41%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/fac2b9c3-e662-d2d7-7b30-cd912b9ca891%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/10fb2ac0-0382-2d7b-13ab-3edf06e50c81%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/de506bc1-37af-0738-1381-56d69e05e829%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/95caf7b5-d139-8b8b-bcd2-874243c4ed50%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/2068586f-bdab-a786-d933-4e32117bb4f8%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/57eb9cac-e684-1a88-3f49-24ed4a7bac47%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/29426284-471b-95bb-be67-a3ca3f9a0d79%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/656c16f7-d561-a8f7-e859-9ac24f357769%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/768de21f-a474-4574-74f4-6d99e3ab542e%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/f20ec6d1-a2b4-9af3-66be-5398dc059c90%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/39bba654-0683-95a4-9092-3c0ecbbf9104%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/0bf6a21a-f667-9851-aeb0-dd6b9b83876e%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/f303adaf-71a3-e855-403d-2a6a3c8f9ceb%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/5acd0507-a96e-7235-e6a5-f31a4c0b7420%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/afae1238-d09f-c934-d363-9b13b733c558%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/b4c5fc80-e197-8d82-ebb0-148675ea7cdd%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddBusinessCard](http://msdn.microsoft.com/library/a30d201b-3073-11c1-0f0c-81c7a3aba6e2%28Office.15%29.aspx)|
|[ClearConversationIndex](http://msdn.microsoft.com/library/5246a0ac-d4e3-4c3b-8362-f5b65e1a28ab%28Office.15%29.aspx)|
|[ClearTaskFlag](http://msdn.microsoft.com/library/833f62c1-2a99-b5ce-76cb-629b195aa63c%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/00a8a4e8-9bdc-d1bc-cb61-c6d925fb754f%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/a9356844-e31e-eb0f-c0f5-a2923ad127db%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/342c6003-e7c5-7314-453c-151fc51d5b2d%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/19ead642-b7bd-579f-e43b-ef5c5d0cfecb%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/5b8c2261-c5ac-fd80-8acf-dfa645a04a1e%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/f2017571-087c-1e83-4003-cb95097d43da%28Office.15%29.aspx)|
|[MarkAsTask](http://msdn.microsoft.com/library/ee38093d-a180-07f7-eae8-c9dbb2e8f413%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/08a0fa20-b891-393a-00fa-5a8fb5405cf6%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/15dc35c1-9dd1-6337-8c61-24d251639d9d%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/c03208a4-dd31-a8ff-0dcd-4ef37a36beb2%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/25a1723a-864b-1526-9897-26e40042f119%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/7d7b5f22-4749-e908-41a7-12a4c730c695%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/b81cf18b-0b0a-19b9-9e88-c6ae0bdc761a%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/78c85013-523e-447b-c47d-2da0705f1fe0%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/212dfd98-c0a2-7f94-249f-ba9baec34882%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/1b7bb1c0-334f-826a-fd6b-8fc3f2fe5d64%28Office.15%29.aspx)|
|[AlternateRecipientAllowed](http://msdn.microsoft.com/library/9ec44a9d-e1e3-ca25-7dc1-a524d1fbfafc%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/d71cb356-f3ae-ab08-4209-1dac0c2b8fdf%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/71f82397-00f3-5660-1211-ebf8b229fff3%28Office.15%29.aspx)|
|[AutoForwarded](http://msdn.microsoft.com/library/822bf508-4a5b-89ec-1077-1cbed75068c2%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/3c0ccbd5-47a6-7a0c-a488-037c48fc1958%28Office.15%29.aspx)|
|[BCC](http://msdn.microsoft.com/library/6454f9b1-1bfa-d4d4-ca95-7a19db920977%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/d1729a7a-5156-bbb5-8a84-347be897af2f%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/578567b1-893b-db4e-dddb-f3c237952c03%28Office.15%29.aspx)|
|[BodyFormat](http://msdn.microsoft.com/library/f635a0bc-20b7-206c-f558-a4ca2519670f%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/049396c0-193b-6c80-9eb0-f55480ffc37a%28Office.15%29.aspx)|
|[CC](http://msdn.microsoft.com/library/c74c1aea-79d1-7096-8f3d-cdd6795fa672%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/7c79286b-13cd-7fb7-c70f-ac12245f9f75%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/1b108d0d-c2b8-60a0-696b-f5c2badd6ead%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/2c93c2a2-4f2f-17af-cba3-91620b3d9c0f%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/97532cd6-397b-303e-b265-7923b371bf9d%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/d97f6416-27c6-b565-9439-a4e9e6f95196%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/d5625f97-3929-95e8-cdaf-6e555cdf9c2b%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/83abef63-4f39-d9dc-9dea-a7365a6461d7%28Office.15%29.aspx)|
|[DeferredDeliveryTime](http://msdn.microsoft.com/library/dbd2fe31-7e5d-d565-61d5-329e8e03b804%28Office.15%29.aspx)|
|[DeleteAfterSubmit](http://msdn.microsoft.com/library/b15d21b5-58d2-4dc2-7244-5e7317f9acd1%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/7d61b284-e3ef-d52c-415c-215206bc5136%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/72ce9938-53fa-ad7c-c69d-453ff348a0e0%28Office.15%29.aspx)|
|[ExpiryTime](http://msdn.microsoft.com/library/18f6497b-6db5-7ec2-7aa8-ec30531e59ef%28Office.15%29.aspx)|
|[FlagRequest](http://msdn.microsoft.com/library/13c04300-ec2a-4ee5-d7b1-eff9f61b71c4%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/06043d0c-c56f-2f87-6018-4a4fa0b0735e%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/9ba8bdbf-1dd5-eaff-3889-33433e3cb3fa%28Office.15%29.aspx)|
|[HTMLBody](http://msdn.microsoft.com/library/c340fe05-9a99-3a32-3d6b-f2f7a568b299%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/77de74c9-e910-e021-1015-6e65f3ead3df%28Office.15%29.aspx)|
|[InternetCodepage](http://msdn.microsoft.com/library/09d80bb8-7677-d9b5-1585-c933af5a7b2d%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/648e6b53-81fb-03ec-0029-edbdd05c663b%28Office.15%29.aspx)|
|[IsMarkedAsTask](http://msdn.microsoft.com/library/6cc4530d-fa74-916b-654d-db995d9a989f%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/620e3af5-0c11-bd78-a98f-b08b36857113%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/91a95fa7-9cbb-0b40-f77f-4f5b3145e0a8%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/7ab16b80-90c6-ef60-b1ce-95fe87ab0d06%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/93194a21-dbec-ebfa-ae5d-d4f287ebb2bd%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/01d9f8bd-d812-7873-02e5-844a64007d5a%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/d8559f9a-b0e5-03ce-febd-e2bd2ca033c9%28Office.15%29.aspx)|
|[OriginatorDeliveryReportRequested](http://msdn.microsoft.com/library/89042dd2-4ac1-109d-5f9c-9ed3733032b0%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/c9328c0e-33d8-4c01-b745-8eb5820a48f5%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/c3ea9b11-9bf2-64c3-409b-3eb33129ae1a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/3aa4d8fe-f6eb-6d09-3475-3d77ca76a9ca%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/394173d4-344a-148a-1628-b4ca47d4ef2d%28Office.15%29.aspx)|
|[PermissionService](http://msdn.microsoft.com/library/c999b215-f360-17b1-4915-45c3b525d3e5%28Office.15%29.aspx)|
|[PermissionTemplateGuid](http://msdn.microsoft.com/library/33436080-1a1c-dee2-5048-83392c241e86%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/bd41eb13-4f66-7de4-8bf7-507ec643be64%28Office.15%29.aspx)|
|[ReadReceiptRequested](http://msdn.microsoft.com/library/5b8d5283-b2fc-4b01-6ccb-b8ac6c7c617e%28Office.15%29.aspx)|
|[ReceivedByEntryID](http://msdn.microsoft.com/library/db4325d3-4442-220d-a812-1d3e4a0085bf%28Office.15%29.aspx)|
|[ReceivedByName](http://msdn.microsoft.com/library/7b57ffcd-b557-f19d-9870-b8c31561120b%28Office.15%29.aspx)|
|[ReceivedOnBehalfOfEntryID](http://msdn.microsoft.com/library/fffcb637-9a7d-3541-49fc-85f314cd92cb%28Office.15%29.aspx)|
|[ReceivedOnBehalfOfName](http://msdn.microsoft.com/library/7a34998b-0475-7279-1e7e-2f0cf2c76bb9%28Office.15%29.aspx)|
|[ReceivedTime](http://msdn.microsoft.com/library/83a4514c-915f-5607-a451-c409720fd25c%28Office.15%29.aspx)|
|[RecipientReassignmentProhibited](http://msdn.microsoft.com/library/f7c7dfbe-d752-c83f-19aa-6eb2f93a85ae%28Office.15%29.aspx)|
|[Recipients](http://msdn.microsoft.com/library/58897f66-8a6a-e1a9-7e3b-5a84624f899d%28Office.15%29.aspx)|
|[ReminderOverrideDefault](http://msdn.microsoft.com/library/78aaca38-6de7-9bc1-6539-74d7b03bfd54%28Office.15%29.aspx)|
|[ReminderPlaySound](http://msdn.microsoft.com/library/7fd10182-445f-2aa6-db9f-2534d66fe0ea%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/f99a945b-1890-7d52-f13b-e0fada91903d%28Office.15%29.aspx)|
|[ReminderSoundFile](http://msdn.microsoft.com/library/11c5ae79-1ce0-5890-1ba1-5a39a88ecc6b%28Office.15%29.aspx)|
|[ReminderTime](http://msdn.microsoft.com/library/ace829f9-a5db-fbce-8948-fde98778d57f%28Office.15%29.aspx)|
|[RemoteStatus](http://msdn.microsoft.com/library/f68f2176-0725-2cdf-572e-3b9f7bea8cb4%28Office.15%29.aspx)|
|[ReplyRecipientNames](http://msdn.microsoft.com/library/96f0e12d-c580-4ec0-9b8f-06607a30faf9%28Office.15%29.aspx)|
|[ReplyRecipients](http://msdn.microsoft.com/library/2d590733-1d67-944e-c2b6-7e08439c1cf5%28Office.15%29.aspx)|
|[RetentionExpirationDate](http://msdn.microsoft.com/library/8f251c3d-8ccc-1378-ad9c-87c6e0ee7d16%28Office.15%29.aspx)|
|[RetentionPolicyName](http://msdn.microsoft.com/library/27e2c3da-ff1a-c261-72cc-b915d89e1019%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/93bfda4f-08fb-9527-6946-625546d7fb49%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/54a436a6-3da4-89d0-e1a6-db45c3732d95%28Office.15%29.aspx)|
|[SaveSentMessageFolder](http://msdn.microsoft.com/library/ab36ae3b-6c6d-842b-dbb4-88c37d8e7874%28Office.15%29.aspx)|
|[Sender](http://msdn.microsoft.com/library/c8afc3f8-fbf5-73b4-43f3-800e18aabb93%28Office.15%29.aspx)|
|[SenderEmailAddress](http://msdn.microsoft.com/library/a157894c-adf2-1cef-ec7c-8516dbef2b7f%28Office.15%29.aspx)|
|[SenderEmailType](http://msdn.microsoft.com/library/e82cb8a6-d480-d1d1-ad15-a498ada6de37%28Office.15%29.aspx)|
|[SenderName](http://msdn.microsoft.com/library/e3c133e6-c7a8-9004-969d-aa2a466f8486%28Office.15%29.aspx)|
|[SendUsingAccount](http://msdn.microsoft.com/library/d4e49128-a63a-d761-90b9-9e1a3305adc7%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/c492be82-093d-547e-85f1-d35c6ee6ba2b%28Office.15%29.aspx)|
|[Sent](http://msdn.microsoft.com/library/a064267f-9329-9018-aa09-c92e17ed46bd%28Office.15%29.aspx)|
|[SentOn](http://msdn.microsoft.com/library/477d7f13-af24-dca7-9845-1a3669093972%28Office.15%29.aspx)|
|[SentOnBehalfOfName](http://msdn.microsoft.com/library/1f58a4b4-abf8-3031-4be1-1538d2d81f5c%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/43272ff5-ab89-f160-7995-981158f6f375%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/10bd56cc-8bdb-470d-a84f-a809c2b057c4%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/5f3e465d-ac2b-a573-0e85-1134e65df017%28Office.15%29.aspx)|
|[Submitted](http://msdn.microsoft.com/library/58dbf39a-962e-8a1d-6424-c66fffeea6d4%28Office.15%29.aspx)|
|[TaskCompletedDate](http://msdn.microsoft.com/library/4bee35d4-1f1e-0b77-2021-84d4916bef8e%28Office.15%29.aspx)|
|[TaskDueDate](http://msdn.microsoft.com/library/161ed0ed-0e3f-2e4c-7e63-daad4e918dd6%28Office.15%29.aspx)|
|[TaskStartDate](http://msdn.microsoft.com/library/76b7109f-55fc-b7e2-63dc-bf7804a709f5%28Office.15%29.aspx)|
|[TaskSubject](http://msdn.microsoft.com/library/f7e4629f-ad47-b455-9fee-b5e537602a34%28Office.15%29.aspx)|
|[To](http://msdn.microsoft.com/library/036dc0b7-1ac7-3884-8d3e-e2f2f1e66ff5%28Office.15%29.aspx)|
|[ToDoTaskOrdinal](http://msdn.microsoft.com/library/d1ccb01a-0792-3779-3f94-eb5195a39bb0%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/af6058cb-abcf-8e77-a5f5-1402addcb333%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/702ae502-d427-eeaf-ddd0-ff9749e7148c%28Office.15%29.aspx)|
|[VotingOptions](http://msdn.microsoft.com/library/696b6dfe-1840-d43b-e6ec-e410a387665c%28Office.15%29.aspx)|
|[VotingResponse](http://msdn.microsoft.com/library/a35c8dd1-57d6-0357-9062-6596a802b8a1%28Office.15%29.aspx)|

## See also


#### Other resources


[Send an E-mail Given the SMTP Address of an Account (Outlook)](http://msdn.microsoft.com/library/5e5f707d-8771-bd5f-945b-58537732d99a%28Office.15%29.aspx)<br>
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
