---
title: TaskItem Object (Outlook)
keywords: vbaol11.chm2990
f1_keywords:
- vbaol11.chm2990
ms.prod: outlook
api_name:
- Outlook.TaskItem
ms.assetid: 5df8cfa5-5460-a5a1-a130-ba5bca1a0091
ms.date: 06/08/2017
---


# TaskItem Object (Outlook)

Represents a task (an assigned, delegated, or self-imposed task to be performed within a specified time frame) in a Tasks folder.


## Remarks

Use the  **[CreateItem](http://msdn.microsoft.com/library/e5fbf367-db16-5042-823e-68e6b805e612%28Office.15%29.aspx)** method to create a **TaskItem** object that represents a new task.

Use  **[Items](http://msdn.microsoft.com/library/441820e7-5fe8-e5ef-83c0-9c87fd3dc9e3%28Office.15%29.aspx)** ( _index_ ), where _index_ is the index number of a task or a value used to match the default property of a task, to return a single **TaskItem** object from a Tasks folder.


## Example

The following Visual Basic for Applications (VBA) example returns a new task.






```
Set myItem = Application.CreateItem(olTaskItem)
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/9d7f10ee-a871-91c3-9c71-309aac23c230%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/989c0e3c-ad11-8017-3b0f-f5e3636c3de6%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/8a0aed80-e92f-a3e8-0341-a55c1a24b6c9%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/2982d79c-81b8-cca9-4a46-ce6b0a95ff80%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/dec504ae-63b3-c668-e81a-cd3ca0cde24c%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/5f0a89ce-b9d7-b7e7-57a5-79a7e69e0d42%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/298eaece-9633-637b-3055-572d77fa3811%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/93d31d5c-fb22-ce19-bcf2-651acc2d5db7%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/6f6acd79-afc2-7b40-60c9-770b8561b1a9%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/390578bf-3c8f-31f1-d81f-e2abba3c1fb6%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/a892d659-1be6-b37e-3a7d-aacf92c19293%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/bee490b1-2ddb-3942-adfe-ed8051b7b0d8%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/b01afdf1-f4a4-8a62-d2c7-bf312ec14f29%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/a2514e35-cdcf-ba93-ad55-b0cc6f64bd78%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/6d093473-9ac3-71a1-9bd6-6511e131afc6%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/b5241171-75d1-17e7-d564-d414662fe5a5%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/93a74a47-b996-5130-74bb-52a662d58a2b%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/001d2598-58e1-86d9-b893-31a79ac2a0a0%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/adc96ece-cea5-c939-7f9a-aa7d0f16960b%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/88e5e300-e036-b511-905c-f0c238c97ade%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/0706a4b9-1035-bdf9-a48d-8d039a2001fa%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/5ec184ae-f512-e38a-0bc0-ddaf519740e2%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/fd96da99-8e7b-249b-7a32-41ac359cb9a6%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/f634105e-5351-6941-e915-ec63cd703b67%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/ff7d2655-06b5-6344-3422-4bf7be761a39%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/29e38bc5-6a19-5144-55ba-207215bd5734%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Assign](http://msdn.microsoft.com/library/f254107a-4182-de3a-2039-08f664e61eeb%28Office.15%29.aspx)|
|[CancelResponseState](http://msdn.microsoft.com/library/564b37c5-f686-8e4d-aa3e-6d41a989b1be%28Office.15%29.aspx)|
|[ClearRecurrencePattern](http://msdn.microsoft.com/library/ad73edd8-d449-5a29-b80f-0717965c40be%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/7682f0c8-d132-2bd6-94e8-6e45fcc00867%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/1224ae94-8c2c-70c8-234a-f3b577cd574e%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/0a2cf917-4899-0fe0-c7dc-35daa70f0892%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/fea0619d-06dc-df44-fe93-5756eefb1be0%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/aa907c9b-b074-fb3b-5134-fd9fa65fa7b9%28Office.15%29.aspx)|
|[GetRecurrencePattern](http://msdn.microsoft.com/library/1937b226-d465-6cc9-7e47-40f4fad1552c%28Office.15%29.aspx)|
|[MarkComplete](http://msdn.microsoft.com/library/e8641735-8bce-6175-d1a7-eb9a69ed8977%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/cc071e73-d165-6082-4016-7ab9d63689d0%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/af648672-6e49-a196-44a2-b9df0b4d3539%28Office.15%29.aspx)|
|[Respond](http://msdn.microsoft.com/library/1befabf7-262f-897a-d1dc-49be4e7ddf9b%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/5b478d20-cd14-2bfa-e96b-0a8d226d451d%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/6f4ae301-089b-047f-bed0-a8faf1583a5a%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/54f751fc-cff1-5d17-f635-f688cd8ad6f8%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/f31b247b-1e8a-6ea8-3d66-cec400e87b70%28Office.15%29.aspx)|
|[SkipRecurrence](http://msdn.microsoft.com/library/19eb8a58-a13f-56ca-b742-a3780d8b0bf1%28Office.15%29.aspx)|
|[StatusReport](http://msdn.microsoft.com/library/70549833-3287-bbbe-6756-896d400f6695%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/896817c2-45f0-afc5-d0a3-bfcdf46b5c2d%28Office.15%29.aspx)|
|[ActualWork](http://msdn.microsoft.com/library/d61075da-bd14-bc59-8f72-b9b675c65f08%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/02b138b9-75ca-04d6-0129-2a5c9917e90c%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/8a645c34-74be-0125-c63f-636c6f090b89%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/19acff0c-a540-f08e-f662-30daf992f575%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/2f777ebd-c53a-f293-9e06-f26234098d12%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/746d3d3d-1b62-0647-60ba-0404d1099926%28Office.15%29.aspx)|
|[CardData](http://msdn.microsoft.com/library/d057d3e6-72c6-01d1-5e1b-37f9ee82cc06%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/c4099fe0-23af-a4cb-dfef-92cbe0c6e600%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/978a3ca8-a444-49ec-593d-370c0deb7710%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/57c71235-ad01-1976-fa5e-2fa2bcfb2d4e%28Office.15%29.aspx)|
|[Complete](http://msdn.microsoft.com/library/c079d11a-bc69-652d-d9c5-6a525f319686%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/50d179df-e029-29d0-9767-2ef441e2305f%28Office.15%29.aspx)|
|[ContactNames](http://msdn.microsoft.com/library/2cbafecb-4984-ed71-efec-c0a565966218%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/69b28ef6-5521-944c-f908-df715e837c36%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/d64f52ce-6657-67bc-a3d6-d2a90155d013%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/ca1eb42a-22b8-8ef9-cf7b-63a96e4910cf%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/0f77fe71-1340-6e50-9de2-fd311e5ae62a%28Office.15%29.aspx)|
|[DateCompleted](http://msdn.microsoft.com/library/17e6a4af-4cd9-0c5e-35ab-5232cf067478%28Office.15%29.aspx)|
|[DelegationState](http://msdn.microsoft.com/library/345321d9-1142-5d6c-dd6a-304b9a4ec4cc%28Office.15%29.aspx)|
|[Delegator](http://msdn.microsoft.com/library/cb0443a3-4ae1-8630-05b9-1b69740163dc%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/65aa9f55-8b53-4c39-e560-c091d397e5ec%28Office.15%29.aspx)|
|[DueDate](http://msdn.microsoft.com/library/4705b840-8bb5-97eb-aa20-1c17cf403653%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/aae660b1-35e5-2cf7-1921-9f91e85d23b1%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/55f086a5-62b3-fbaa-4e7d-de3e0528634b%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/2a2faad7-1030-cdd8-8a8d-8018aad3b667%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/cab606ff-7b3c-4d94-779d-c8b07a5913ab%28Office.15%29.aspx)|
|[InternetCodepage](http://msdn.microsoft.com/library/a9186d58-a6b3-8269-56ab-105456883283%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/de713a49-bdc8-363e-4990-cf3535b27981%28Office.15%29.aspx)|
|[IsRecurring](http://msdn.microsoft.com/library/09684a02-bab4-56ff-cdb3-0a20049c968d%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/c22c45f7-eb0d-457b-359e-6a3833d1fcfe%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/d59018c1-2d32-9081-9e65-8a4627c62ab4%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/0dd93a32-1857-1304-b52d-1deb282984ea%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/e5deb86e-ad13-32f0-8dd8-802e7cc539aa%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/3cc676b5-4817-adab-9a72-61a0214a2f64%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/3cc48820-8c03-57ab-6c7f-d4b47aed9fbe%28Office.15%29.aspx)|
|[Ordinal](http://msdn.microsoft.com/library/533ad2a0-a46b-2fbb-dc5f-29a2b838fe83%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/5b7a31be-0c9f-b8f3-7cc3-0c117aa0f809%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/14ce6d04-10fb-a1e1-75a1-82b19ea76f9e%28Office.15%29.aspx)|
|[Owner](http://msdn.microsoft.com/library/8af59077-9f4f-2099-fd98-416061447968%28Office.15%29.aspx)|
|[Ownership](http://msdn.microsoft.com/library/7eb09c39-77af-6522-8194-a8369a577342%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/9fdcce5a-d094-dccd-5081-edbabdd2fb5a%28Office.15%29.aspx)|
|[PercentComplete](http://msdn.microsoft.com/library/39525055-647b-02c0-a9da-150698181511%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/f6fc4753-5eee-8892-4cd3-3df74b2fce18%28Office.15%29.aspx)|
|[Recipients](http://msdn.microsoft.com/library/03743284-9753-6cb9-b5cc-20bc5cb3621e%28Office.15%29.aspx)|
|[ReminderOverrideDefault](http://msdn.microsoft.com/library/3a11ee36-3418-422e-0783-e39bf92ded6f%28Office.15%29.aspx)|
|[ReminderPlaySound](http://msdn.microsoft.com/library/22698193-bc36-c2fb-3ee1-d04d1e3a15a6%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/441de5fb-2c43-9024-b4cb-126f683df9f5%28Office.15%29.aspx)|
|[ReminderSoundFile](http://msdn.microsoft.com/library/29bfa689-08b6-f963-9ecb-3744b1032062%28Office.15%29.aspx)|
|[ReminderTime](http://msdn.microsoft.com/library/c9a0526f-a986-76df-80e2-f085fd645df8%28Office.15%29.aspx)|
|[ResponseState](http://msdn.microsoft.com/library/91f1d4a1-f55b-7379-c1a8-c302bac25a6c%28Office.15%29.aspx)|
|[Role](http://msdn.microsoft.com/library/122d18ee-6a60-4a40-1b3e-66b8bd1c8a9d%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/ff94ab2c-7e34-0eb5-3aeb-b7805b5e9a2c%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/65ed9320-7c1f-4201-3b13-30fa0df9381b%28Office.15%29.aspx)|
|[SchedulePlusPriority](http://msdn.microsoft.com/library/773a9232-692f-cc5c-795f-2f36466afaf4%28Office.15%29.aspx)|
|[SendUsingAccount](http://msdn.microsoft.com/library/711382c3-1003-cf0e-2f29-fc3f9d4320a8%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/b4b91017-bae5-4766-37ec-606cf57683e5%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/f2c0a916-b654-98de-c134-d9736d482cea%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/9949591e-987a-12e5-0ba0-01a078c7e7e4%28Office.15%29.aspx)|
|[StartDate](http://msdn.microsoft.com/library/0ec17958-78cd-3a2e-05c3-cbc8e367e3df%28Office.15%29.aspx)|
|[Status](http://msdn.microsoft.com/library/fc575f57-0651-f620-89df-3bbaa89e019d%28Office.15%29.aspx)|
|[StatusOnCompletionRecipients](http://msdn.microsoft.com/library/9800dcb7-6b12-af4b-0379-25658c946118%28Office.15%29.aspx)|
|[StatusUpdateRecipients](http://msdn.microsoft.com/library/904e4685-75db-9267-7f88-dd2bce6e8509%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/9f487fbc-48ab-e01d-c1a4-5b67fcb1a118%28Office.15%29.aspx)|
|[TeamTask](http://msdn.microsoft.com/library/a405ff6d-0061-5fd4-e3a7-9550c9d12e1f%28Office.15%29.aspx)|
|[ToDoTaskOrdinal](http://msdn.microsoft.com/library/dae1be0d-aef7-2901-2c23-8014434e5d8c%28Office.15%29.aspx)|
|[TotalWork](http://msdn.microsoft.com/library/3b940a69-f2b4-30d1-0027-49450f547b01%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/98ce2f8f-37f8-ab98-7cd6-2e70550d1805%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/d4dce54f-412b-c4b4-4553-3f8df9551ac0%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[TaskItem Object Members](http://msdn.microsoft.com/library/97234a76-2fc5-bbe4-2e14-25ae18694fc9%28Office.15%29.aspx)
