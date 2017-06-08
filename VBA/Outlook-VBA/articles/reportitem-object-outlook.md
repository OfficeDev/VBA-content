---
title: ReportItem Object (Outlook)
keywords: vbaol11.chm3007
f1_keywords:
- vbaol11.chm3007
ms.prod: outlook
api_name:
- Outlook.ReportItem
ms.assetid: 16ebe336-72e0-42f6-99d3-edecc3ea284d
ms.date: 06/08/2017
---


# ReportItem Object (Outlook)

Represents a mail-delivery report in an Inbox folder. 


## Remarks

The  **ReportItem** object is similar to a **[MailItem](http://msdn.microsoft.com/library/14197346-05d2-0250-fa4c-4a6b07daf25f%28Office.15%29.aspx)** object, and it contains a report (usually the non-delivery report) or error message from the mail transport system.

Unlike other Microsoft Outlook objects, you cannot create this object. Report items are created automatically when any report or error in general is received from the mail transport system.


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/a585b4f0-9453-da34-6360-f7cb72943af9%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/e57a3f9b-f5a5-e345-aca7-1ab0a1c141e3%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/16c7acf4-015e-b9ab-bd72-a54921de8709%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/9df574ed-f1df-2ff8-1508-4d2ab35a8bca%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/c8b45b3b-627c-4851-b743-2612828546b0%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/105baaa6-b0ff-d7dc-6181-b8c9141c192b%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/65377c41-b51a-779c-9892-a61cc6e9b9da%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/3fa6311c-e7d3-3a08-f416-05c4c718a916%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/c4bfb8ad-3fa2-2319-fd83-5784aa4ab203%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/c3a2882c-ff82-39a1-3d18-5bf4f608b09e%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/a1d1a844-96c0-50f0-0db8-d0f6980d422d%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/2fca7e89-39b3-73c4-715a-003921a055cd%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/dc485dac-3ee0-f20e-c9b8-6dd01b56ac30%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/d20e50a8-c73d-d866-0cd0-d6085a3b6eb6%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/33212db2-878f-1672-1fc9-90ddd4800f0c%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/8b75f239-a3c2-01fc-1b94-84b2b680a420%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/607369d8-5e04-f9c8-ad11-828e185edef2%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/f44fe7fe-29b3-f1ab-70ee-0e395ad6896a%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/5fd89535-8fa4-202e-bb0a-1dc4d608acec%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/7b142bcb-dd96-a0ec-5684-b7311f34d772%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/f73cb164-0c88-f439-6474-a4502b6731ea%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/e2f835e3-9f25-8cbb-3ba7-5b0e7e495c63%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/b5724798-8c73-13ce-23d4-9d7ec8147f44%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/aab0b0f3-8e33-f1fa-cc74-d914effcb833%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/934c4793-0809-65dc-4805-de28a54634cf%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/1656ff7c-85c9-f193-3312-279d35622008%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Close](http://msdn.microsoft.com/library/bd38dde1-b747-5686-6073-1945557c9926%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/f667600e-ca34-b8a9-9c3d-3b598888dfe3%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/1a206718-6ba6-6b1f-803e-93b1ee435dc0%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/583673cd-f646-2843-82ce-b11d673df5a3%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/1e8d3031-1a14-25b0-997f-ef27c42e2e61%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/171a46e4-bd39-9556-36f3-0c0c60ed2b0b%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/48d486f5-dd1f-2e82-017e-6c14aace4d1b%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/cfe23d31-8cf7-afc0-3232-b59e55e4a30b%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/70497e98-0b4d-266b-10c1-c340a14e82c9%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/d73cf745-580c-47c9-c011-55d88460295e%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/ac2959ca-7ac0-c308-060b-6a273fade806%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/d827ed53-ce2e-c8cf-485e-970125d03045%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/d7d93015-1d16-c217-cbc0-5e866c1ba89b%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/55f74600-8058-b7cc-33c3-e5b80cef255a%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/3241eac3-d93f-3686-2f2d-5619c967b7c2%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/f10b5de0-1b2b-b401-b5fd-4486ed2fd4ed%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/57983279-5be9-1a08-8a13-d70d5e252699%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/241a6cf7-6b53-fece-907c-455c979d2405%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/293e2355-5597-2628-8eaa-8e2504fc8510%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/9f5740ed-e740-17bc-f073-a3e551466113%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/b642a06e-94f0-b615-1806-fdd5ae881d48%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/c70ebc07-c07d-963c-b757-01035ded7be9%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/021d0822-d4a3-ec4a-eb27-b64bc2deaac1%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/5c7665b6-fb36-8e5e-4f90-6997fa108fd3%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/e81a4cc1-b94f-b5cb-7224-68d90c075f8b%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/00dc7cb0-aa06-1e08-74c8-3cb5e3540a03%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/f296b505-28c6-ee81-0ad3-72a5ad611f9e%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/2a9ec97b-56c5-f93c-eb42-7ddb93a4697e%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/4ecffe39-45d5-c646-2de2-50bf440189c7%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/ec5db93a-43e5-8f9c-ed55-c940c0d056d1%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/ec1ea335-6ccd-2b9e-398b-f4b44d017c41%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/85f457b7-b344-30cd-de7c-b1dfd1a7ee6d%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/6abf44c2-975d-90ca-986f-f1d8b7c1ba6b%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/096bfebc-20eb-ea36-cff8-a96a514b5903%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/bd3b3dfe-6368-6ba7-c609-8b0e3ea97a27%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/5f693704-0e16-4a45-2136-b7aa945003b2%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/a8c61bf4-b9d3-fefd-dbe2-37d9ac7c36cc%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/85b79531-6475-5403-8974-0c3cf836018b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/b8663e30-f169-9050-a5ab-cf8573053e40%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/090bedce-4517-1d8c-9c46-1f67bcced7fa%28Office.15%29.aspx)|
|[RetentionExpirationDate](http://msdn.microsoft.com/library/96ddf279-5cfe-0245-302d-816d3f020e39%28Office.15%29.aspx)|
|[RetentionPolicyName](http://msdn.microsoft.com/library/054e4a80-a00e-62c1-f442-50d5340eb36e%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/64b02e02-9d33-da89-5293-276c1f3eb3cb%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/a5d225a9-5667-43df-a580-8c20cf69438a%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/b9599afe-1c2b-36b2-2ce4-8e781f32975a%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/4554eed6-44a8-7f88-63f2-f06de1e8694c%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/0c4ed1df-3ebd-3b0c-2ea7-548cc6576481%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/368ee0bb-4167-2499-4a83-4b4a4320eae0%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/a42224a1-ab82-7533-2c75-882f99f49e8b%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[ReportItem Object Members](http://msdn.microsoft.com/library/5a5662dd-e969-bbd5-129b-44609ba1cf9f%28Office.15%29.aspx)
