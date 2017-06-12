---
title: AppointmentItem Object (Outlook)
keywords: vbaol11.chm2988
f1_keywords:
- vbaol11.chm2988
ms.prod: outlook
api_name:
- Outlook.AppointmentItem
ms.assetid: 204a409d-654e-27aa-643a-8344c631b82d
ms.date: 06/08/2017
---


# AppointmentItem Object (Outlook)

Represents a meeting, a one-time appointment, or a recurring appointment or meeting in the Calendar folder.


## Remarks

Use the  **[CreateItem](http://msdn.microsoft.com/library/e5fbf367-db16-5042-823e-68e6b805e612%28Office.15%29.aspx)** method to create an **AppointmentItem** object that represents a new appointment.

Use  **[Items](http://msdn.microsoft.com/library/89a031e0-c0a3-fc22-f485-189df8db45f4%28Office.15%29.aspx)** ( _index_ ), where _index_ is the index number of an appointment or a value used to match the default property of an appointment, to return a single **AppointmentItem** object from a Calendar folder.

You can also return an  **AppointmentItem** object from a **[MeetingItem](meetingitem-object-outlook.md)** object by using the **[GetAssociatedAppointment](http://msdn.microsoft.com/library/8344d40d-5c1d-ead3-87cb-fd795b831712%28Office.15%29.aspx)** method.

When you work with recurring appointment items, you should release any prior references, obtain new references to the recurring appointment item before you access or modify the item, and release these references as soon as you are finished and have saved the changes. This practice applies to the recurring  **AppointmentItem** object, and any **[Exception](http://msdn.microsoft.com/library/010552b0-9ba6-c81b-1e3a-fd6a681e5163%28Office.15%29.aspx)** or **[RecurrencePattern](recurrencepattern-object-outlook.md)** object. To release a reference in Visual Basic for Applications (VBA) or Visual Basic, set that existing object to **Nothing**. In C#, explicitly release the memory for that object.

Note that even after you release your reference and attempt to obtain a new reference, if there is still an active reference, held by another add-in or Outlook, to one of the above objects, your new reference will still point to an out-of-date copy of the object. Therefore, it is important that you release your references as soon as you are finished with the recurring appointment.

The following code example in VBA shows how to release and refresh references in order to obtain up-to-date data for a recurring appointment. The example obtains a set of appointment items from the Calendar folder. It assumes that the first item in the appointment collection is part of a recurring appointment. The example shows that a reference to the appointment collection obtained before an exception is created does not reflect the exception. The example then releases this reference and other existing appointment references, after which new references that point to the appointment collection reflect the exception.




```
Sub TestExceptions() 
 
 Dim oItems As Items 
 
 Dim oItemOriginal As AppointmentItem 
 
 Dim oItemNew As AppointmentItem 
 
 Dim rPattern As RecurrencePattern 
 
 Dim oEx As Exceptions 
 
 Dim oEx2 As Exceptions 
 
 Dim oOccurrence As AppointmentItem 
 
 Dim i As Long 
 
 
 
 ' This is the initial reference to an appointment collection. 
 
 Set oItems = _ 
 
 Outlook.Application.Session.GetDefaultFolder(olFolderCalendar).Items 
 
 
 
 ' This is the original reference to the first appointment in the 
 
 ' collection before an exception is created. 
 
 Set oItemOriginal = oItems.Item(1) 
 
 
 
 ' Code example assumes that the first appointment in the collection 
 
 ' is a recurring appointment. 
 
 Set oOccurrence = _ 
 
 oItemOriginal.GetRecurrencePattern().GetOccurrence(#2/28/2010 8:00:00 AM#) 
 
 
 
 ' Create an exception by changing the 2/28 occurrence to 3/3. 
 
 oOccurrence.Start = #3/3/2010 8:00:00 AM# 
 
 oOccurrence.Save 
 
 
 
 Stop 
 
 
 
 ' Preexisting reference to the first appointment in the collection 
 
 ' does not reflect the exception. 
 
 oItemOriginal.Save 
 
 Set oEx = oItemOriginal.GetRecurrencePattern().Exceptions 
 
 Debug.Print oItemOriginal.subject 
 
 Debug.Print " Original item exceptions: " &amp; oEx.Count 
 
 
 
 ' Get a new reference based on the existing reference to the 
 
 ' appointment collection created before the exception. 
 
 ' The new reference does not reflect the exception. 
 
 Set oItemNew = oItems.Item(1) 
 
 oItemNew.Save 
 
 Set oEx2 = oItemNew.GetRecurrencePattern().Exceptions 
 
 Debug.Print " New item exceptions: " &amp; oEx2.Count 
 
 
 
 ' Same: preexisting reference to the first appointment in the collection 
 
 ' does not reflect the exception. 
 
 Set oEx = oItemOriginal.GetRecurrencePattern().Exceptions 
 
 Debug.Print " Original item exceptions: " &amp; oEx.Count 
 
 
 
 ' Release all existing references to appointment items, 
 
 ' including the appointment collection, an exception, occurrence, 
 
 ' or any other appointment. 
 
 Debug.Print "REFRESH ITEM COLLECTION" 
 
 Set oItems = Nothing 
 
 Set oItemNew = Nothing 
 
 Set oEx = Nothing 
 
 Set oEx2 = Nothing 
 
 Set oOccurrence = Nothing 
 
 Set oItemOriginal = Nothing 
 
 Set rPattern = Nothing 
 
 
 
 ' Get new references to appointment items, including the appointment 
 
 ' collection, individual appointments, and exceptions. 
 
 Set oItems = _ 
 
 Outlook.Application.Session.GetDefaultFolder(olFolderCalendar).Items 
 
 Set oItemNew = oItems.Item(1) 
 
 
 
 ' If no other add-ins have the same recurring appointment open, 
 
 ' the new references reflect the current exception count. 
 
 Set oEx2 = oItemNew.GetRecurrencePattern().Exceptions 
 
 Debug.Print " New item exceptions: " &amp; oEx2.Count 
 
 
 
 Debug.Print "RE-GET ORIGINAL" 
 
 Set oItemOriginal = oItems.Item(1) 
 
 Set oEx = oItemOriginal.GetRecurrencePattern().Exceptions 
 
 Debug.Print " Original item exceptions: " &amp; oEx.Count 
 
End Sub
```


## Example

The following Visual Basic for Applications (VBA) example returns a new appointment.


```
Set myItem = Application.CreateItem(olAppointmentItem)
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/e63e0d48-f3cd-4c3b-1ef9-4a9a83a34a32%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/4b048018-99af-22b8-66b5-1f876856c6a8%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/69ea3ad0-5cb6-a832-8e46-9ed86c59c3b2%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/db3778f6-3cf4-0830-909a-0b3171b6d605%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/1574e5b0-b2d1-ca0a-3e1a-0c5efef0a99c%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/8a78394e-cc83-f965-4c28-83c574282c44%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/db38a11b-c9bc-ebda-5900-00391cdf080f%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/1add142b-e23a-adb5-66b9-184be82087a1%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/7754a2f9-d36b-5ba8-331c-8dfcfa9f03d3%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/c24e39d1-39e5-6422-78ff-9d4e391ea2ae%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/e68833b3-c585-725a-aa71-bbba9ffbad16%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/dc6944f6-e020-bdd7-0b64-98a3f3d2e94c%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/c5a696e6-96c3-ac4f-d81b-e103b8c091c5%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/1af9cc71-36d1-e759-5151-401c899ae13b%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/bd16129d-d9e3-2953-2ccb-116eadd5bbaa%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/f40abc41-efb5-d36e-229b-0b9fbbcf63cd%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/3d56ee04-9a9a-1f10-0436-a2b567b17229%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/08a0d07b-6fd0-690e-6745-f5ad92bb3ff1%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/82bb6104-ce62-8fb6-1472-d84fd36e94ac%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/aa39ec06-19ed-4655-6990-e4c4c45649d5%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/749e8d58-c15c-0b63-5486-cc2aa2190638%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/bc3ea8eb-15eb-ef97-e292-e74799cce150%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/c49245b9-0770-f551-c382-3f5745aead04%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/6571ae2f-4964-f38f-e39e-14a2b94caa73%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/9629cf4d-99e7-c751-0543-15daf41df49c%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/55539ad2-d53e-b28e-06f4-13c5f545a89b%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[ClearRecurrencePattern](http://msdn.microsoft.com/library/a880839a-7c0a-7940-95f7-ee3699e88ece%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/61072885-5319-5a00-c4f1-d412eb4d60cb%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/947f1cfd-f60c-a47e-ba4d-3ffde8c13c91%28Office.15%29.aspx)|
|[CopyTo](http://msdn.microsoft.com/library/50b8e820-fdb9-1ee9-289b-99be037300c4%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/5114b1ca-d923-9de2-cbad-8b14be001deb%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/24706404-d646-a3ac-b7b1-64a6a1c697a9%28Office.15%29.aspx)|
|[ForwardAsVcal](http://msdn.microsoft.com/library/5d5456b4-315c-b9e3-2ed8-a1b709999a2e%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/661386aa-c357-8437-36a4-c0a910573b90%28Office.15%29.aspx)|
|[GetOrganizer](http://msdn.microsoft.com/library/c6cd89b6-d0ab-721b-5671-c057b0646c24%28Office.15%29.aspx)|
|[GetRecurrencePattern](http://msdn.microsoft.com/library/a9f67c5b-a77f-4e34-e654-d12560a6dba0%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/29f3a845-cf7d-e598-45c5-1e67e8985215%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/8052ae60-f7eb-e932-7ec4-176262727351%28Office.15%29.aspx)|
|[Respond](http://msdn.microsoft.com/library/060d1fcb-0011-bea0-5c6b-fa3538ff9a2d%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/177980e8-96cc-a72e-ede3-7aad3a98cf68%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/24dc2663-ca45-395d-5c7f-6a6eaaff120f%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/72f2e997-55ef-b98b-fdd1-7f3b810a28ed%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/5b79f252-ffce-a59d-873f-48efe467df3b%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/8c2c91c4-b242-df8d-a8d1-b6493cf95bdd%28Office.15%29.aspx)|
|[AllDayEvent](http://msdn.microsoft.com/library/42803963-dce2-9eb1-bddb-62867abc57b5%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/a40911dd-9513-8d55-03b7-1aa52b81e24d%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/4d2eb321-84c7-5613-35cc-9df3e872541d%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/d48a7ba9-bb70-9126-98ef-3bdee1f62436%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/f6d1c066-dfda-0267-e4b9-ca65345bcc6e%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/0b9cec33-d3bf-1902-cc60-36966c06b757%28Office.15%29.aspx)|
|[BusyStatus](http://msdn.microsoft.com/library/38a07f42-121d-86a4-68fe-0c508ddb265a%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/71360959-7a42-7aa8-579f-1e544a734dd0%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/8955081b-3868-ea81-f136-3948fc49f219%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/79ed2563-1dc8-a6c5-d715-a11940cb9176%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/1c32306c-1852-8eab-a924-1b0f7e59dd58%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/6897e23d-1d1d-f8fb-fbab-aa19242f4e7f%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/10748cab-d404-019e-1eaa-9aa8102a1ce0%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/dc46a62a-2259-80a8-3abf-ce214d9c911b%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/d1bb179b-5ac4-d1e8-0b49-bca0e2ec1f16%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/0b380863-817d-0f5e-0117-464ab218dbb2%28Office.15%29.aspx)|
|[Duration](http://msdn.microsoft.com/library/eea64bdd-c19b-01c7-4fdb-111df86de2c4%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/ce40f8ef-224e-2a64-fe78-cf4ae42be822%28Office.15%29.aspx)|
|[EndInEndTimeZone](http://msdn.microsoft.com/library/9fec38c1-3cd1-d428-4d51-48e01954ee03%28Office.15%29.aspx)|
|[EndTimeZone](http://msdn.microsoft.com/library/8f33d93f-c0fe-fda1-608d-dec7fb86c732%28Office.15%29.aspx)|
|[EndUTC](http://msdn.microsoft.com/library/c741e893-3a29-10cc-0730-a0796d8c2e4c%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/8f4160de-0840-902a-589e-bce80797b6f5%28Office.15%29.aspx)|
|[ForceUpdateToAllAttendees](http://msdn.microsoft.com/library/fe926820-2694-9aa3-8359-cc2ed3ac2f32%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/516c9628-54e5-4732-9845-f359804dff64%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/6d0dc447-80f3-ab00-4bb9-7bbda34745aa%28Office.15%29.aspx)|
|[GlobalAppointmentID](http://msdn.microsoft.com/library/3a5e210a-5298-8977-d6e4-dc49a59bdd78%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/72b5a262-b7d0-4fca-5862-5d932cf9c694%28Office.15%29.aspx)|
|[InternetCodepage](http://msdn.microsoft.com/library/7ebb4076-7ba0-cae4-f6d4-e85d37675a8e%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/d0c14fa2-6bfe-29e8-e68b-3eff01a8bd70%28Office.15%29.aspx)|
|[IsRecurring](http://msdn.microsoft.com/library/93e243cc-fec9-2474-6828-5077bfd744e7%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/4fac93ef-e927-9751-10f2-297e1b054c2b%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/723d152c-cd71-6038-1eed-06de4c96c32c%28Office.15%29.aspx)|
|[Location](http://msdn.microsoft.com/library/bde4d455-15de-bb29-c27e-99c34836bd46%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/63fe552b-2721-2a9c-5fec-ad0d666065b6%28Office.15%29.aspx)|
|[MeetingStatus](http://msdn.microsoft.com/library/cfd970cd-df6c-4537-0a17-b5adab3b667f%28Office.15%29.aspx)|
|[MeetingWorkspaceURL](http://msdn.microsoft.com/library/f4b6708b-70ab-d20c-4c28-c6d0d2d991f0%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/e98318d9-72e9-0914-83c6-3a05f544874f%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/4562097b-3489-768c-f808-84249e030aab%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/f2e1844f-638a-259e-4ed6-e814da9a1330%28Office.15%29.aspx)|
|[OptionalAttendees](http://msdn.microsoft.com/library/019262e6-34cd-8138-0237-13e7b99e51d7%28Office.15%29.aspx)|
|[Organizer](http://msdn.microsoft.com/library/20fac1d5-0d40-918d-909d-a86069e6ed1d%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/2c7777c1-8195-db3e-0ca6-c2eeeb42f23c%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/25ba176e-5525-dd25-25d5-523de4c91d3b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/5bc7bcec-18bb-ebfb-a8e4-329a354841cd%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/096e33b6-cb70-765c-c469-955ae7c7d840%28Office.15%29.aspx)|
|[Recipients](http://msdn.microsoft.com/library/4fc824fb-b046-558c-7aa7-28586cd11a7d%28Office.15%29.aspx)|
|[RecurrenceState](http://msdn.microsoft.com/library/dd435d09-8cb0-8efe-c947-88c90951f64e%28Office.15%29.aspx)|
|[ReminderMinutesBeforeStart](http://msdn.microsoft.com/library/d83269fc-b706-d285-d8ec-23fed4952955%28Office.15%29.aspx)|
|[ReminderOverrideDefault](http://msdn.microsoft.com/library/08c20608-6065-1e4a-8c89-8aa35c682c77%28Office.15%29.aspx)|
|[ReminderPlaySound](http://msdn.microsoft.com/library/4020684b-c89d-7371-75e0-4f3dfe01bec3%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/575d5fb2-1672-ddae-832c-7dcc7d1da2d6%28Office.15%29.aspx)|
|[ReminderSoundFile](http://msdn.microsoft.com/library/e3599e63-1300-7821-b94d-f8387a47e87d%28Office.15%29.aspx)|
|[ReplyTime](http://msdn.microsoft.com/library/cf455d15-a360-818a-b6a7-59f4d1e89f4c%28Office.15%29.aspx)|
|[RequiredAttendees](http://msdn.microsoft.com/library/8ff112e9-2d8c-89de-0bdf-e8b9998f9269%28Office.15%29.aspx)|
|[Resources](http://msdn.microsoft.com/library/9b989d76-6897-cd2d-9156-fd7391dad8c1%28Office.15%29.aspx)|
|[ResponseRequested](http://msdn.microsoft.com/library/a96727b8-1a8a-6ab6-d8a0-4ca9c9409aff%28Office.15%29.aspx)|
|[ResponseStatus](http://msdn.microsoft.com/library/853cf25d-6cfc-baef-b906-acf43dbd6478%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/12af0270-e9bc-88ce-1d36-eafadf698406%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/b009f0a8-fbd4-23f7-01fd-72faf73d0bd0%28Office.15%29.aspx)|
|[SendUsingAccount](http://msdn.microsoft.com/library/c3a73b32-c2e1-cb32-35e3-e278f78700ad%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/1e5aec44-3328-f6fe-6ee4-019a4afc8d21%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/ff92a5eb-5a5a-9211-c247-42b9d993780f%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/72dd6cfd-67a1-23d6-df95-174becd97f03%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/1b869a9d-fe08-6efb-48b1-f33cf9ea0024%28Office.15%29.aspx)|
|[StartInStartTimeZone](http://msdn.microsoft.com/library/4735816e-2c3b-816c-434d-8d7ea42fec81%28Office.15%29.aspx)|
|[StartTimeZone](http://msdn.microsoft.com/library/3259fa91-5f6c-b899-9bfc-2ac669911271%28Office.15%29.aspx)|
|[StartUTC](http://msdn.microsoft.com/library/8bfbf95f-bd88-acdc-f592-c41b454afe4b%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/57f0f242-6d04-175f-4ea2-25145787f5bd%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/297e8b98-54b6-bd05-31e2-8479ae06ceb3%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/d1245b91-62e9-78b8-9012-85c11959166c%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[How to: Import Appointment XML Data into Outlook Appointment Objects](http://msdn.microsoft.com/library/ecfd3849-877b-01ad-2b76-1a54e980f6e2%28Office.15%29.aspx)
[AppointmentItem Object Members](http://msdn.microsoft.com/library/c72c459d-6d3c-7a05-aa4a-b1b767ddc0b2%28Office.15%29.aspx)
