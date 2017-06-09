---
title: MeetingItem Object (Outlook)
keywords: vbaol11.chm2989
f1_keywords:
- vbaol11.chm2989
ms.prod: outlook
api_name:
- Outlook.MeetingItem
ms.assetid: b75730f5-b395-3d66-5acd-b64fd8fcd78f
ms.date: 06/08/2017
---


# MeetingItem Object (Outlook)

Represents a change to the recipient's Calendar folder initiated by another party or as a result of a group action.


## Remarks

Unlike other Microsoft Outlook objects, you cannot create this object. It is created automatically when you set the  **[MeetingStatus](http://msdn.microsoft.com/library/cfd970cd-df6c-4537-0a17-b5adab3b667f%28Office.15%29.aspx)** property of an **[AppointmentItem](appointmentitem-object-outlook.md)** object to **olMeeting** and send it to one or more users. They receive it in their inboxes as a **MeetingItem**.

Use the  **[GetAssociatedAppointment](http://msdn.microsoft.com/library/8344d40d-5c1d-ead3-87cb-fd795b831712%28Office.15%29.aspx)** method to return the **AppointmentItem** object associated with a **MeetingItem** object, and work directly with the **AppointmentItem** object to respond to the request.


## Example

The following example uses the  **[CreateItem](http://msdn.microsoft.com/library/e5fbf367-db16-5042-823e-68e6b805e612%28Office.15%29.aspx)** method to create an appointment. It becomes a **MeetingItem** with both a required and an optional attendee when it is received in the inbox of each of the recipients.


```
Set myItem = myOlApp.CreateItem(olAppointmentItem) 
 
myItem.MeetingStatus = olMeeting 
 
myItem.Subject = "Strategy Meeting" 
 
myItem.Location = "Conference Room B" 
 
myItem.Start = #9/24/97 1:30:00 PM# 
 
myItem.Duration = 90 
 
Set myRequiredAttendee = myItem.Recipients.Add("Nate _ 
 
 Sun") 
 
myRequiredAttendee.Type = olRequired 
 
Set myOptionalAttendee = myItem.Recipients.Add("Kevin _ 
 
 Kennedy") 
 
myOptionalAttendee.Type = olOptional 
 
Set myResourceAttendee = _ 
 
 myItem.Recipients.Add("Conference Room B") 
 
myResourceAttendee.Type = olResource 
 
myItem.Send
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/10fab1af-e29f-74d2-5fae-aa61822f06dd%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/ea34a56f-abdc-c928-9df8-ba83d3584565%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/50ec1cf8-98cc-390b-0080-74d6e145524d%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/418fcee8-fba8-1296-0689-75d4f84c508a%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/9550ed34-0e04-eee0-b149-4df496c8e155%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/4b52c888-fd21-478b-d396-915f7c5a193e%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/17ffaaa1-fe71-d21c-e4cf-884321f9afe2%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/1ed68d13-6368-05f4-99ad-c7db8997eb34%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/26bbc5fc-4a65-101b-9693-f8d9ed9421c9%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/59de272e-a36a-e842-a962-03ebe2befa26%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/451d1b1b-3411-1f0a-69f7-14a1fc9071d9%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/155c5225-aeb0-55b6-26dc-811d00128238%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/da5383b0-c2bd-d0b2-b023-c493d469d3d2%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/9af94b62-d992-39e8-ddce-507db6a2febb%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/c9ba1402-f1e1-3bb6-3242-288cd0276224%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/b3d05c13-4b5d-032b-49bb-18c4f4a626b5%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/0d12864b-07ca-5f97-8aab-ea9415e8b44c%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/d286705a-d542-f3aa-3121-f0635e0cc62c%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/6bc3629b-b08a-0d8b-f1e3-6d3c90176ac2%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/8a83b213-1afb-7ded-eb67-3e5d21502c5b%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/17ef8085-38ac-7e32-7704-54a2f2224e87%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/5b1ffaf2-f2ad-081a-423c-85c16a38e68b%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/d93bd51d-a169-0007-4188-4fff829dbb1e%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/9dc87c39-d209-dc06-86e8-ce00f9cb152f%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/87053a2f-11cc-6a76-a4fd-7c752efb00bd%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/22a52e41-cbc5-ced7-a942-ae06035aebbb%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Close](http://msdn.microsoft.com/library/f88f72a4-9fec-8576-191f-4f800f0e0929%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/a79ddac2-c1ef-76e2-9baa-446e4a4d6e98%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/62821244-206b-039d-d321-e1b373a44d0b%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/8b6f7748-7a96-0ab2-c11f-3c7e9b729b05%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/ca456d91-43db-3f94-133b-913fd50ef4bc%28Office.15%29.aspx)|
|[GetAssociatedAppointment](http://msdn.microsoft.com/library/8344d40d-5c1d-ead3-87cb-fd795b831712%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/0ff1d250-a791-4438-4b3a-112b76a18ea8%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/ab888dbc-f31f-ac68-f914-c97d6af2e6d9%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/fe53eccd-cd6b-ecf5-2fa4-c56de616686d%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/df43c9d0-8a70-a54a-90a2-9675414ccddb%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/b3a85859-dd31-d1ca-8ce5-d8a2b06576bb%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/51af858c-18d7-ea94-5b0b-27ad45037fc4%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/cda4cccc-1930-3aa8-d0e1-651de6b0a0b7%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/d9a6ea8c-2146-06ec-aa8b-6e39fd60a916%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/e4530fc8-2e6b-ad84-936c-9d20c4c0bff2%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/f659ae7a-27bb-5be3-ef8b-2dd07e8bcdf2%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/d727319a-3e42-c053-6ee7-550d13dfc738%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/7399ae31-606a-816a-6049-7bd5778b829b%28Office.15%29.aspx)|
|[AutoForwarded](http://msdn.microsoft.com/library/30fe7984-771b-146f-ae16-5bee257457f1%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/5a6c9fbb-0f41-9b69-dd41-35ec72e16c7c%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/be9dc49d-c6f6-736d-afee-f44661f98823%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/e8e92565-d86d-8306-3281-cefa42f5ffd6%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/ae4a9569-afb6-a7d7-2cbb-351141f99588%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/92750064-bb55-ba4b-83c3-b3d74da5ea50%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/cf9ddbc6-286d-47ba-8fb2-6e54d70fc302%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/8cdf2d98-8780-1fac-cc11-4e36f93aab29%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/67a28933-1f89-8f1d-9217-bacd61aa85b9%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/0c1ab025-e215-57fb-78ff-6260d45e6ad9%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/89b43466-1ac3-3323-235f-2231ae6656b6%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/aa20cc5a-6c28-858d-dc3f-1d5c8b30013c%28Office.15%29.aspx)|
|[DeferredDeliveryTime](http://msdn.microsoft.com/library/1d68f55d-dd1c-f043-8d7b-f96f0e981cbc%28Office.15%29.aspx)|
|[DeleteAfterSubmit](http://msdn.microsoft.com/library/576ca136-8144-dc32-e048-d75a17303740%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/bd5afbb2-570f-6d0c-5108-20119839f43e%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/03d2684e-9608-f631-284d-ed63ce11c85a%28Office.15%29.aspx)|
|[ExpiryTime](http://msdn.microsoft.com/library/14e78315-f430-20fe-b24e-fe8bf396bc3b%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/700ddd9d-8cc8-9fbd-1520-24e0257c4dae%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/5e170a6a-6857-ca24-4c14-1e2bc046fd2d%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/f8dd738d-efd5-730d-f976-2f582b932db2%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/1e84c838-06f6-823f-1605-8085d42bb0a0%28Office.15%29.aspx)|
|[IsLatestVersion](http://msdn.microsoft.com/library/aee3a832-b1b5-538d-dd45-e64769662dfc%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/b15a928d-8e49-0303-0fe2-e2debbe228ec%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/175726cb-b1fa-83ab-8e14-684611fab02b%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/a5a0bc64-4129-f93e-ff07-2a1785a10099%28Office.15%29.aspx)|
|[MeetingWorkspaceURL](http://msdn.microsoft.com/library/ad97f3cc-35c6-b653-73b9-7c7a0555afe2%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/0e7f893f-4de3-06c6-32e0-c815f9af35d5%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/42bcb344-a9d5-bb3e-f346-d41cc1f30055%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/e4365923-032c-6dfc-a79e-1b2c63b417b8%28Office.15%29.aspx)|
|[OriginatorDeliveryReportRequested](http://msdn.microsoft.com/library/7dfa8dfe-0268-57d8-0ba2-7f69789d4ce9%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/77e6ed76-e562-2b0b-d0ca-65675afa842a%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/48f8c948-9fbd-842a-e9c0-5eb021e283e7%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/174f14b5-8c30-ae21-21fe-0672a4b2de06%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/63e35352-ec63-c7cb-2e94-eb8022cff8a9%28Office.15%29.aspx)|
|[ReceivedTime](http://msdn.microsoft.com/library/bf27c544-3f3e-87e1-9f0c-84f1469d771d%28Office.15%29.aspx)|
|[Recipients](http://msdn.microsoft.com/library/486f7f16-1db9-b99e-d5b0-0e94edc7a745%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/a86c2e82-061b-a608-ca22-1a4a8973a62e%28Office.15%29.aspx)|
|[ReminderTime](http://msdn.microsoft.com/library/c34a0a59-79f6-e1ee-7e69-762e6a6de731%28Office.15%29.aspx)|
|[ReplyRecipients](http://msdn.microsoft.com/library/a4314327-6174-4fb2-236a-e154457033ae%28Office.15%29.aspx)|
|[RetentionExpirationDate](http://msdn.microsoft.com/library/81ce85c5-0b0e-40b0-563a-8654cd3dece4%28Office.15%29.aspx)|
|[RetentionPolicyName](http://msdn.microsoft.com/library/a17f0781-f290-a2f8-10a9-af75b51e9a1f%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/4bf67ee1-f0bc-92b8-948f-2de7807a1dd3%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/911ad89a-15f3-ce02-0eba-4081b43b0e72%28Office.15%29.aspx)|
|[SaveSentMessageFolder](http://msdn.microsoft.com/library/35c8c917-0ae6-f2ac-dd34-79a62cc321f3%28Office.15%29.aspx)|
|[SenderEmailAddress](http://msdn.microsoft.com/library/b318c074-4897-d99d-2b7c-870b4ab083e9%28Office.15%29.aspx)|
|[SenderEmailType](http://msdn.microsoft.com/library/99870104-54f2-cce5-ff32-212bd335a4c5%28Office.15%29.aspx)|
|[SenderName](http://msdn.microsoft.com/library/07dd4ff2-36cd-cfbd-3b48-08e60f0aed78%28Office.15%29.aspx)|
|[SendUsingAccount](http://msdn.microsoft.com/library/81713c7b-dfb0-eb91-b017-82b427bee823%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/5f7dfd4d-d51f-9bd2-5125-0fab980f3509%28Office.15%29.aspx)|
|[Sent](http://msdn.microsoft.com/library/b95be57b-8332-3423-4438-c84a8612bc7c%28Office.15%29.aspx)|
|[SentOn](http://msdn.microsoft.com/library/361dfa26-6514-cc3a-aa1b-240728ac0dd9%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/b18a448d-c3a6-e8cd-f251-30883e53e484%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/8c19d83c-0b75-2760-6808-3fd8cea3e4b9%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/f390f25e-2dfa-f4f9-a9af-3d694de241c9%28Office.15%29.aspx)|
|[Submitted](http://msdn.microsoft.com/library/195c6188-eaab-3319-0b69-641d273b406f%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/5d556f3d-96bd-fa20-cc96-37c98150079a%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/a88bfccb-e90b-1327-29e4-afb63565bb1b%28Office.15%29.aspx)|

## See also


#### Other resources


[MeetingItem Object Members](http://msdn.microsoft.com/library/9ae6a19d-d326-4c37-90d8-5ed9933672a0%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
