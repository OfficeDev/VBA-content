---
title: ContactItem Object (Outlook)
keywords: vbaol11.chm2992
f1_keywords:
- vbaol11.chm2992
ms.prod: outlook
api_name:
- Outlook.ContactItem
ms.assetid: 8e32093c-a678-f1fd-3f35-c2d8994d166f
ms.date: 06/08/2017
---


# ContactItem Object (Outlook)

Represents a contact in a Contacts folder.


## Remarks

A contact can represent any person with whom you have any personal or professional contact.

Use the  **[CreateItem](http://msdn.microsoft.com/library/e5fbf367-db16-5042-823e-68e6b805e612%28Office.15%29.aspx)** method to create a **ContactItem** object that represents a new contact.

Use  **[Items](http://msdn.microsoft.com/library/441820e7-5fe8-e5ef-83c0-9c87fd3dc9e3%28Office.15%29.aspx)** ( _index_ ), where _index_ is the index number of a contact or a value used to match the default property of a contact, to return a single **ContactItem** object from a Contacts folder.


## Example

The following Visual Basic for Applications (VBA) example returns a new contact.


```
Set myItem = Application.CreateItem(olContactItem)
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/d771b7ab-9235-2b62-60df-f4a168ba75e2%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/ef818f33-7ed8-7beb-1fb8-83eb01c271a5%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/5c240669-e37d-12ea-7094-e070884907e8%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/e7080603-d978-aeb8-a50c-1bcc53504422%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/d0c0bfd1-5d18-759c-0131-c78e45982b18%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/7451778c-801a-15a9-203d-1a1c61ebc155%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/ba862dea-f2e1-a864-f6c3-a8987c28bfcf%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/c4c33ade-25db-f9d9-69fb-97dcce76bf45%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/d6e84398-10ca-53fc-8576-102ae8d8971f%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/c9fe9c4d-3c00-455c-3e89-9ac584597117%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/7ad6f4cd-d993-2c5b-ebce-8a3561c39a54%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/a37ddcea-12eb-82f8-19a7-609d599394b2%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/cebd1e59-b3a4-3c9d-5ed1-ff95c2c3d1ed%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/beeeb53c-94fe-ae1b-7870-87bd37b3debf%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/e2f6da0c-0470-8cbd-ce31-2e2a6e0e5353%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/46112f35-cbca-6bf6-3c4a-28be9013007c%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/d09448bb-09de-03be-4f4b-98f3a94bce6c%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/80f12bd2-a36d-d5ae-e6a1-55df6fe2fc2c%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/4138deee-2915-f581-b003-16007e37f128%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/508b4637-9d74-7645-7719-3c148d0688d8%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/1700ad85-3113-e937-9eb3-be78246fd4d5%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/0560988f-95a1-23f5-67af-f94321d9ff39%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/380f187f-e914-5810-baaf-07473f1719f1%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/28c7171e-df79-8a5d-5c3c-138ec3b3ee9b%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/16a3d7ce-0843-5eb5-bbea-df6557ceda05%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/934a4bac-8b75-246b-97ed-214ebd3fbd8f%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddBusinessCardLogoPicture](http://msdn.microsoft.com/library/73e19806-6892-f378-cc38-70e9d90922d1%28Office.15%29.aspx)|
|[AddPicture](http://msdn.microsoft.com/library/aa02c3b2-bb4c-fde9-3dbf-f871cbc200b1%28Office.15%29.aspx)|
|[ClearTaskFlag](http://msdn.microsoft.com/library/19e4fecd-7565-60ae-707b-410f4c1a6dcd%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/17cd04b5-1bf1-5df1-b1f4-f6e488d00fd5%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/0e99dbcb-95f0-b1a2-e709-165a09035354%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/229d4c37-4659-01ae-0623-3e1095b13048%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/789611b5-7079-2290-738f-64266cedbe2a%28Office.15%29.aspx)|
|[ForwardAsBusinessCard](http://msdn.microsoft.com/library/2f1a74c3-86f0-a054-75e2-272dbb261fb7%28Office.15%29.aspx)|
|[ForwardAsVcard](http://msdn.microsoft.com/library/3d4f0154-9860-823f-c316-c88e410b59c3%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/19609cbf-d6ad-8a66-5a42-0010cd2797ee%28Office.15%29.aspx)|
|[MarkAsTask](http://msdn.microsoft.com/library/def25d8d-6074-5e4d-18d9-82381b0b7876%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/e5e2ac9f-5fb2-2ebb-4afe-b61fc414d0aa%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/97546d46-1171-76f6-1f4f-e02cc39110a3%28Office.15%29.aspx)|
|[RemovePicture](http://msdn.microsoft.com/library/a67d9d39-1697-0780-b52f-a3cc463f60d9%28Office.15%29.aspx)|
|[ResetBusinessCard](http://msdn.microsoft.com/library/a6eed85a-ac25-64c6-6bf3-650d5129c8e3%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/1f7e998f-be59-6a50-95b5-cb066adbb278%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/9f563508-e7fc-ee35-366b-6937604cf25f%28Office.15%29.aspx)|
|[SaveBusinessCardImage](http://msdn.microsoft.com/library/889728f2-2c17-6b83-a858-bb32ef5845e6%28Office.15%29.aspx)|
|[ShowBusinessCardEditor](http://msdn.microsoft.com/library/96db2b87-02b2-f97e-cff4-9d852fc875d6%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/22613243-1281-82b7-5da3-da1f4d620599%28Office.15%29.aspx)|
|[ShowCheckAddressDialog](http://msdn.microsoft.com/library/773a1a3c-1247-fd48-399a-728766e56570%28Office.15%29.aspx)|
|[ShowCheckFullNameDialog](http://msdn.microsoft.com/library/d42632e3-6f50-cce7-80c6-cf846be1f925%28Office.15%29.aspx)|
|[ShowCheckPhoneDialog](http://msdn.microsoft.com/library/3ef93046-c2b0-5707-9bb1-4dbfb5d7366c%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Account](http://msdn.microsoft.com/library/0d75eabd-f0f8-538d-576d-c75a0b41c84a%28Office.15%29.aspx)|
|[Actions](http://msdn.microsoft.com/library/1fd1e1ad-d5ab-75ab-eb73-c5521d5801a7%28Office.15%29.aspx)|
|[Anniversary](http://msdn.microsoft.com/library/c1e9a355-9776-0baa-90b6-743cea99b4e6%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/ab4f247f-a0e3-fd33-90dd-c961b792cb1d%28Office.15%29.aspx)|
|[AssistantName](http://msdn.microsoft.com/library/0695875e-fbeb-3786-ca58-bb56644b2fff%28Office.15%29.aspx)|
|[AssistantTelephoneNumber](http://msdn.microsoft.com/library/0dcb4d55-1dbf-0fca-d1a4-ef5af715fc52%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/5679948f-bb5b-661a-0060-7941a8e436ef%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/f14ae270-0d3d-5b8c-c85c-9809ba0b82fa%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/c41719c5-0f26-aa0a-754c-c72127c88e00%28Office.15%29.aspx)|
|[Birthday](http://msdn.microsoft.com/library/d36f2719-8ccb-a6bf-457c-7430e9c26853%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/5da750b7-90c2-a46b-99e9-0365340b53fa%28Office.15%29.aspx)|
|[Business2TelephoneNumber](http://msdn.microsoft.com/library/ba436db9-61e1-5913-4209-efec732c652e%28Office.15%29.aspx)|
|[BusinessAddress](http://msdn.microsoft.com/library/840e40ed-6773-3ef0-d17a-471921415bf9%28Office.15%29.aspx)|
|[BusinessAddressCity](http://msdn.microsoft.com/library/6c21e0f0-ab9b-5190-6749-4e8f6fc909e8%28Office.15%29.aspx)|
|[BusinessAddressCountry](http://msdn.microsoft.com/library/cd5b1640-ddbd-9fca-062c-f03ed39f7821%28Office.15%29.aspx)|
|[BusinessAddressPostalCode](http://msdn.microsoft.com/library/0c9f643a-c29e-4ae5-cea7-f54b3e98b543%28Office.15%29.aspx)|
|[BusinessAddressPostOfficeBox](http://msdn.microsoft.com/library/447b3e5d-7f8f-372f-d5a6-843ba65a72b7%28Office.15%29.aspx)|
|[BusinessAddressState](http://msdn.microsoft.com/library/0d8d9136-6d41-b0ed-f320-6e26fca15cf7%28Office.15%29.aspx)|
|[BusinessAddressStreet](http://msdn.microsoft.com/library/1d3e67c4-b02d-c2cf-b04b-85bc1464d788%28Office.15%29.aspx)|
|[BusinessCardLayoutXml](http://msdn.microsoft.com/library/0a2cfc55-7835-db1a-7dba-b896e14a13d5%28Office.15%29.aspx)|
|[BusinessCardType](http://msdn.microsoft.com/library/57de9454-83e0-976f-cb69-d472bfd9fb3c%28Office.15%29.aspx)|
|[BusinessFaxNumber](http://msdn.microsoft.com/library/85468b34-1ad3-ecec-92ee-af6ca68616be%28Office.15%29.aspx)|
|[BusinessHomePage](http://msdn.microsoft.com/library/96ef88dd-be24-17f1-1584-8db35747a088%28Office.15%29.aspx)|
|[BusinessTelephoneNumber](http://msdn.microsoft.com/library/6c30e792-f5d6-bdd3-1b01-ca9a5bf27b44%28Office.15%29.aspx)|
|[CallbackTelephoneNumber](http://msdn.microsoft.com/library/2750a396-a88d-c7f2-a353-ab7de2715339%28Office.15%29.aspx)|
|[CarTelephoneNumber](http://msdn.microsoft.com/library/45538c71-eacd-603a-4325-f3e3f3b2c523%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/c2ac3005-caa9-cc91-766e-a341ed0d0e9e%28Office.15%29.aspx)|
|[Children](http://msdn.microsoft.com/library/e002308f-4488-ad1f-a6de-3768c8c2f414%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/7c08cb72-fdbb-aac8-2691-382bfdae22c8%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/38fb0e7a-a5e6-6f3f-5c59-0cdc4a4af53f%28Office.15%29.aspx)|
|[CompanyAndFullName](http://msdn.microsoft.com/library/99a9087d-c511-f274-f506-b07a26cb9050%28Office.15%29.aspx)|
|[CompanyLastFirstNoSpace](http://msdn.microsoft.com/library/dd8b1ac3-b671-c1a3-bbc3-8c2cdeefaaca%28Office.15%29.aspx)|
|[CompanyLastFirstSpaceOnly](http://msdn.microsoft.com/library/8f78b5c8-3832-8c30-6ba6-d7f0149d2dd3%28Office.15%29.aspx)|
|[CompanyMainTelephoneNumber](http://msdn.microsoft.com/library/21e092ae-d0cf-fc6c-6834-f0db032409d5%28Office.15%29.aspx)|
|[CompanyName](http://msdn.microsoft.com/library/076cd6f7-7faa-ab1c-254c-3307c40520ee%28Office.15%29.aspx)|
|[ComputerNetworkName](http://msdn.microsoft.com/library/3042c37b-08b5-25d6-f83d-f038789f844a%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/c51d7028-40d5-4d67-7bc6-8715bfa89f24%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/13a4e7cf-66b3-fba6-b179-68eaf1de8db6%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/696feb03-5fda-3abc-8633-0b096298dafe%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/b22f1484-b24b-db16-96ae-1cf49c0f89ed%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/95c48449-2488-4110-a498-e9a9a563d54f%28Office.15%29.aspx)|
|[CustomerID](http://msdn.microsoft.com/library/863c6dec-2375-7e7b-45bf-69fcc920b948%28Office.15%29.aspx)|
|[Department](http://msdn.microsoft.com/library/661beecc-f6aa-7215-ba01-b075209f2ad3%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/3067720e-dea5-f14f-0c46-61184078fd4f%28Office.15%29.aspx)|
|[Email1Address](http://msdn.microsoft.com/library/0bd407bc-21a9-16e6-709d-383cb79b4d6e%28Office.15%29.aspx)|
|[Email1AddressType](http://msdn.microsoft.com/library/f498f1be-713c-7d86-28c8-fbeb6b1d3f6d%28Office.15%29.aspx)|
|[Email1DisplayName](http://msdn.microsoft.com/library/71a7e227-f462-9dae-1315-dfe445c2329c%28Office.15%29.aspx)|
|[Email1EntryID](http://msdn.microsoft.com/library/8329e2a9-52e6-f3f1-56b4-c17752510e0b%28Office.15%29.aspx)|
|[Email2Address](http://msdn.microsoft.com/library/1656eb41-55b3-50f7-7351-b287e07bcac0%28Office.15%29.aspx)|
|[Email2AddressType](http://msdn.microsoft.com/library/09e1448e-87d7-5040-a13f-ae8d7ae67cb9%28Office.15%29.aspx)|
|[Email2DisplayName](http://msdn.microsoft.com/library/37a4cbfb-8318-d968-353d-bee87536794e%28Office.15%29.aspx)|
|[Email2EntryID](http://msdn.microsoft.com/library/0c5691bb-e112-763b-d126-2bcc2c52ccce%28Office.15%29.aspx)|
|[Email3Address](http://msdn.microsoft.com/library/b0f29077-a06c-a2cf-e873-b9d560d91498%28Office.15%29.aspx)|
|[Email3AddressType](http://msdn.microsoft.com/library/af814290-2f74-5d83-28b0-a0af055a63cc%28Office.15%29.aspx)|
|[Email3DisplayName](http://msdn.microsoft.com/library/ea4dc96f-dd92-213e-cda0-b6a619e8965c%28Office.15%29.aspx)|
|[Email3EntryID](http://msdn.microsoft.com/library/f38c8002-c4a8-f47a-c783-986e4121f4c3%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/04f4bd28-5edf-4e69-5b7c-d3bec749fc4f%28Office.15%29.aspx)|
|[FileAs](http://msdn.microsoft.com/library/65ea6d2e-17ef-5de8-adfb-2b1aebfbd9fd%28Office.15%29.aspx)|
|[FirstName](http://msdn.microsoft.com/library/403b5e5a-037b-cf21-efc2-2bd2a80c3789%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/618b1bda-696c-9232-f68b-37613940ab20%28Office.15%29.aspx)|
|[FTPSite](http://msdn.microsoft.com/library/24f6f207-763f-5a5b-83f1-ba099a780b67%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/3036dc57-31fb-45ad-f51e-49336206581d%28Office.15%29.aspx)|
|[FullNameAndCompany](http://msdn.microsoft.com/library/931d6e82-4d0a-7d6e-8c30-7f64d783884e%28Office.15%29.aspx)|
|[Gender](http://msdn.microsoft.com/library/0192a64e-d575-d43f-77ed-adbcc156786f%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/d1f8530f-f797-413f-92cb-d0e8215de0e4%28Office.15%29.aspx)|
|[GovernmentIDNumber](http://msdn.microsoft.com/library/cfe46380-7b96-441f-f111-e7c796ed6bab%28Office.15%29.aspx)|
|[HasPicture](http://msdn.microsoft.com/library/5e835af9-bcee-692d-f486-5f8a4a0efa1a%28Office.15%29.aspx)|
|[Hobby](http://msdn.microsoft.com/library/6386f34c-ac9c-cd81-75ec-01ac00c75f8b%28Office.15%29.aspx)|
|[Home2TelephoneNumber](http://msdn.microsoft.com/library/18a3b191-e27d-7459-82aa-1138fbacbb21%28Office.15%29.aspx)|
|[HomeAddress](http://msdn.microsoft.com/library/c7ba836b-4b55-cedb-35f6-e6540bdf2c58%28Office.15%29.aspx)|
|[HomeAddressCity](http://msdn.microsoft.com/library/1d2334f2-0401-3bcc-53bf-fa55e1664d9c%28Office.15%29.aspx)|
|[HomeAddressCountry](http://msdn.microsoft.com/library/a3e1f178-c01c-e7df-ee4e-fc82f89915f0%28Office.15%29.aspx)|
|[HomeAddressPostalCode](http://msdn.microsoft.com/library/28d65f71-6be6-5d9e-0935-7f09a5f9fa94%28Office.15%29.aspx)|
|[HomeAddressPostOfficeBox](http://msdn.microsoft.com/library/9c1b310d-13d8-407c-a97e-a52405e37fb2%28Office.15%29.aspx)|
|[HomeAddressState](http://msdn.microsoft.com/library/bc052902-1e38-3d6a-1b7b-308861357731%28Office.15%29.aspx)|
|[HomeAddressStreet](http://msdn.microsoft.com/library/9a7af500-e817-6fb1-89b4-6b0ef70741bf%28Office.15%29.aspx)|
|[HomeFaxNumber](http://msdn.microsoft.com/library/ee7c8d16-4cdf-8b98-dc76-b7d9d8f64f07%28Office.15%29.aspx)|
|[HomeTelephoneNumber](http://msdn.microsoft.com/library/d8e6ffa0-2d1b-384a-070f-2511be2a7a90%28Office.15%29.aspx)|
|[IMAddress](http://msdn.microsoft.com/library/d7f916b0-aa5b-872d-0928-bbab5000ac75%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/f56f1c98-3d07-87d5-2af2-c98ef314036f%28Office.15%29.aspx)|
|[Initials](http://msdn.microsoft.com/library/f1daa747-1c53-f244-6a08-cd6147a02ff3%28Office.15%29.aspx)|
|[InternetFreeBusyAddress](http://msdn.microsoft.com/library/b45fdf0f-1474-5a67-b628-f74e3c1aabb8%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/35ff3a52-2d2a-458f-3e16-4a8f674bb0fa%28Office.15%29.aspx)|
|[ISDNNumber](http://msdn.microsoft.com/library/98e27ef6-0af7-948c-8f62-49bc01d42c11%28Office.15%29.aspx)|
|[IsMarkedAsTask](http://msdn.microsoft.com/library/bf651a37-e486-1c54-83d4-3bb3714f7187%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/44d2bc7e-30f1-0b74-f9e2-0e3db5d6212a%28Office.15%29.aspx)|
|[JobTitle](http://msdn.microsoft.com/library/6a08691c-7747-d9de-2349-5a3fbb01b136%28Office.15%29.aspx)|
|[Journal](http://msdn.microsoft.com/library/3916e2e9-9660-6622-6315-cf1a21865f53%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/68f14566-71b5-24ae-5b9d-e8337b97ed78%28Office.15%29.aspx)|
|[LastFirstAndSuffix](http://msdn.microsoft.com/library/b234614c-e2c0-cba2-6ec8-69be1a31caf1%28Office.15%29.aspx)|
|[LastFirstNoSpace](http://msdn.microsoft.com/library/2ddd5572-453c-970f-b6d6-5831a394a5cc%28Office.15%29.aspx)|
|[LastFirstNoSpaceAndSuffix](http://msdn.microsoft.com/library/15c9527b-3837-d4a0-0249-2cd751e4379f%28Office.15%29.aspx)|
|[LastFirstNoSpaceCompany](http://msdn.microsoft.com/library/52e60375-954d-ff0d-d06e-9b0fe8823184%28Office.15%29.aspx)|
|[LastFirstSpaceOnly](http://msdn.microsoft.com/library/ab1e1edc-23af-ceaf-64e7-d8604c689752%28Office.15%29.aspx)|
|[LastFirstSpaceOnlyCompany](http://msdn.microsoft.com/library/93f08c59-78d5-d007-98a5-dfb940d1e84a%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/14962429-bbf6-a814-647a-70be1dad339d%28Office.15%29.aspx)|
|[LastName](http://msdn.microsoft.com/library/430682f6-a230-887b-404b-a71989121fa2%28Office.15%29.aspx)|
|[LastNameAndFirstName](http://msdn.microsoft.com/library/7667650d-3da9-8a30-63d5-2d6b0d55ccb7%28Office.15%29.aspx)|
|[MailingAddress](http://msdn.microsoft.com/library/7af2770c-1f8b-510b-4e6f-3ef919082088%28Office.15%29.aspx)|
|[MailingAddressCity](http://msdn.microsoft.com/library/f9b8510a-998a-bf7e-9fa5-f567f9d784bc%28Office.15%29.aspx)|
|[MailingAddressCountry](http://msdn.microsoft.com/library/0c6aaaa2-7d09-0c65-cbf6-4c1413095ecd%28Office.15%29.aspx)|
|[MailingAddressPostalCode](http://msdn.microsoft.com/library/bdb1cd44-1ae5-598d-0f25-604deafdb7ed%28Office.15%29.aspx)|
|[MailingAddressPostOfficeBox](http://msdn.microsoft.com/library/b4dc4baa-2af8-f008-6f26-3070dd739a6c%28Office.15%29.aspx)|
|[MailingAddressState](http://msdn.microsoft.com/library/9e15bba8-2256-fd1a-60ae-ac63d6d4f4e3%28Office.15%29.aspx)|
|[MailingAddressStreet](http://msdn.microsoft.com/library/8487bbf4-0d48-4224-9370-e4e78f100d09%28Office.15%29.aspx)|
|[ManagerName](http://msdn.microsoft.com/library/bf8c6303-75da-f589-c7a0-b16ded036bb3%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/8d5f49e4-7941-47f7-e6f1-b2ddc145d0d4%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/3d6594b7-8abe-9e49-64e0-be3062807e34%28Office.15%29.aspx)|
|[MiddleName](http://msdn.microsoft.com/library/07e0c9b1-1093-2f8a-3b89-ba8570b2bdf5%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/8c88b294-8c43-386c-36c4-749377862938%28Office.15%29.aspx)|
|[MobileTelephoneNumber](http://msdn.microsoft.com/library/425023bb-b7c6-628f-7c23-ac3dc1adb5ec%28Office.15%29.aspx)|
|[NetMeetingAlias](http://msdn.microsoft.com/library/ee7b35bb-7006-04f3-c98e-93d393630532%28Office.15%29.aspx)|
|[NetMeetingServer](http://msdn.microsoft.com/library/884d7542-c2df-2f55-5000-4bbf05849418%28Office.15%29.aspx)|
|[NickName](http://msdn.microsoft.com/library/d970aad5-0197-8cf5-b6f1-8d768734d785%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/d1d68995-31f0-de56-7711-d414c970ca00%28Office.15%29.aspx)|
|[OfficeLocation](http://msdn.microsoft.com/library/faf658b0-61ff-26ec-4a65-09dfd564f9a4%28Office.15%29.aspx)|
|[OrganizationalIDNumber](http://msdn.microsoft.com/library/3d77cd1e-7688-8410-8766-c88ec56ed3da%28Office.15%29.aspx)|
|[OtherAddress](http://msdn.microsoft.com/library/16bc351b-9522-4cf9-2838-74e644fec828%28Office.15%29.aspx)|
|[OtherAddressCity](http://msdn.microsoft.com/library/ab29f816-1434-658b-196b-a918a4234aa7%28Office.15%29.aspx)|
|[OtherAddressCountry](http://msdn.microsoft.com/library/c9fd6c5f-db32-e1d6-1f2f-88c0c12285c7%28Office.15%29.aspx)|
|[OtherAddressPostalCode](http://msdn.microsoft.com/library/a9cecb5e-d6c3-9496-8537-fab14520321f%28Office.15%29.aspx)|
|[OtherAddressPostOfficeBox](http://msdn.microsoft.com/library/905500a2-475a-ed2a-79b5-e46a3d8c117c%28Office.15%29.aspx)|
|[OtherAddressState](http://msdn.microsoft.com/library/a8073ae6-eb63-5674-16c1-ceb83babda25%28Office.15%29.aspx)|
|[OtherAddressStreet](http://msdn.microsoft.com/library/dd82de5e-63fc-18bb-5211-f8218e08354b%28Office.15%29.aspx)|
|[OtherFaxNumber](http://msdn.microsoft.com/library/9e0d701e-874f-6cd8-dae5-4b7a0b5f5744%28Office.15%29.aspx)|
|[OtherTelephoneNumber](http://msdn.microsoft.com/library/21a2f846-64ea-0898-dc37-4fe6dbe9ab49%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/44511a6c-8be6-8897-90b5-76d56da5b7ca%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/235a541d-2482-b3ec-af37-aec9150500f7%28Office.15%29.aspx)|
|[PagerNumber](http://msdn.microsoft.com/library/2b83aa60-4766-64bc-f590-6f58ba631c32%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/4aa19d6d-c15d-c7ac-731c-7a2d886665d2%28Office.15%29.aspx)|
|[PersonalHomePage](http://msdn.microsoft.com/library/cbc6abda-eb66-acfd-20db-f5572d20d602%28Office.15%29.aspx)|
|[PrimaryTelephoneNumber](http://msdn.microsoft.com/library/be4fb227-597f-99ba-09b1-fdc4dbd5f60a%28Office.15%29.aspx)|
|[Profession](http://msdn.microsoft.com/library/4aeadd8a-d227-7a51-ba01-c67fd94ed3a3%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/e69b37ce-1d3c-7cef-591c-83e12c76733c%28Office.15%29.aspx)|
|[RadioTelephoneNumber](http://msdn.microsoft.com/library/130631d8-6b1b-1378-2937-ced00ec5c70d%28Office.15%29.aspx)|
|[ReferredBy](http://msdn.microsoft.com/library/052e1595-dd0f-d240-712d-e460bf78a1bf%28Office.15%29.aspx)|
|[ReminderOverrideDefault](http://msdn.microsoft.com/library/08e77dff-b325-c565-746a-e47e4d66ed77%28Office.15%29.aspx)|
|[ReminderPlaySound](http://msdn.microsoft.com/library/a9941154-6c65-57c7-1dab-6d6a59620d92%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/8e3b7091-1d4c-8d9a-ebb0-ebe478c6e386%28Office.15%29.aspx)|
|[ReminderSoundFile](http://msdn.microsoft.com/library/aafbdc5b-816f-3605-d265-5da349e9e791%28Office.15%29.aspx)|
|[ReminderTime](http://msdn.microsoft.com/library/c8b62f1b-693d-65fc-863d-df407571a7e4%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/f8e7e632-113b-a50e-211b-dbd182221168%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/eecebb77-163a-de3c-26b8-8a5916749e18%28Office.15%29.aspx)|
|[SelectedMailingAddress](http://msdn.microsoft.com/library/7f0a68a0-2663-276f-7217-f580d63edb51%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/f0b31e8d-573f-242a-63f4-09b0d86b54a4%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/b67eb0d4-9b97-2be7-fc24-ecdd58fb01ca%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/34f42cff-f7be-815b-6165-c9e58b586e4a%28Office.15%29.aspx)|
|[Spouse](http://msdn.microsoft.com/library/4ca95e03-ec75-702a-3d7a-f2f36822d3b7%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/99c06ab3-1ecd-062f-0b47-1b102b136cbb%28Office.15%29.aspx)|
|[Suffix](http://msdn.microsoft.com/library/edb92ed2-c42d-9f0d-b67a-e58ccd72ea0f%28Office.15%29.aspx)|
|[TaskCompletedDate](http://msdn.microsoft.com/library/6567575d-f95f-b409-a298-a19a590ff1d7%28Office.15%29.aspx)|
|[TaskDueDate](http://msdn.microsoft.com/library/3449ec3e-ca65-c8e3-c3fc-ca9eb5ab0f75%28Office.15%29.aspx)|
|[TaskStartDate](http://msdn.microsoft.com/library/f84e949f-4126-39e9-b0b9-e27e5ef3951a%28Office.15%29.aspx)|
|[TaskSubject](http://msdn.microsoft.com/library/f80c702f-70fa-d7c4-fcc5-b85d802a8d40%28Office.15%29.aspx)|
|[TelexNumber](http://msdn.microsoft.com/library/f20ec303-71fa-982d-5b69-384ef666f19c%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/3dd64517-ccb1-fed3-4f90-c407fc09d5e4%28Office.15%29.aspx)|
|[ToDoTaskOrdinal](http://msdn.microsoft.com/library/080e32ad-b770-42d1-60d0-4eb6271056db%28Office.15%29.aspx)|
|[TTYTDDTelephoneNumber](http://msdn.microsoft.com/library/88d6c5d6-c6cb-c873-8ef2-c3293c1fd81a%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/6029ff4d-76be-d0df-a5b4-c7af42f2fa17%28Office.15%29.aspx)|
|[User1](http://msdn.microsoft.com/library/eae210af-eca7-8229-d2a3-eaca2c357f6c%28Office.15%29.aspx)|
|[User2](http://msdn.microsoft.com/library/6155ee5e-076a-2560-a220-e0dd07e243ba%28Office.15%29.aspx)|
|[User3](http://msdn.microsoft.com/library/feac1ac5-9598-7183-7262-6f28e23efaaa%28Office.15%29.aspx)|
|[User4](http://msdn.microsoft.com/library/9146bfe5-4abc-c335-3dc9-11427583c792%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/f52b8fb8-945b-a406-b3cb-1c9dcc150184%28Office.15%29.aspx)|
|[WebPage](http://msdn.microsoft.com/library/0914b59d-64f3-2c6f-fc83-25d5f0e91abb%28Office.15%29.aspx)|
|[YomiCompanyName](http://msdn.microsoft.com/library/23316fb2-4211-6b1e-4ead-dadcb35965dd%28Office.15%29.aspx)|
|[YomiFirstName](http://msdn.microsoft.com/library/aa69a838-692d-f9bc-4c39-b561121f7125%28Office.15%29.aspx)|
|[YomiLastName](http://msdn.microsoft.com/library/42f21ac7-cca2-a8b1-88b7-012b0bc3f0c9%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[ContactItem Object Members](http://msdn.microsoft.com/library/a8b13369-4c87-02aa-e62a-1f3067e559fa%28Office.15%29.aspx)
