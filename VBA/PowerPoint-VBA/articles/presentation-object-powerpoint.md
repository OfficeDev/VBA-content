---
title: Presentation Object (PowerPoint)
keywords: vbapp10.chm524000
f1_keywords:
- vbapp10.chm524000
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation
ms.assetid: ec75cf52-69f8-d35b-0a26-4a8da8a9683f
ms.date: 06/08/2017
---


# Presentation Object (PowerPoint)

Represents a Microsoft PowerPoint presentation. 


## Remarks

The  **Presentation** object is a member of the **[Presentations](presentations-object-powerpoint.md)** collection. The **Presentations** collection contains all the **Presentation** objects that represent open presentations in PowerPoint.

The following examples describe how to:


- Return a presentation that you specify by name or index number
    
- Return the presentation in the active window
    
- Return the presentation in any document window or slide show window you specify
    

## Example

Use  **Presentations** (index), where index is the presentation's name or index number, to return a single **Presentation** object. The name of the presentation is the file name, with or without the file name extension, and without the path. The following example adds a slide to the beginning of Sample Presentation.


```
Presentations("Sample Presentation").Slides.Add 1, 1
```

Note that if multiple presentations with the same name are open, the first presentation in the collection with the specified name is returned.

Use the [ActivePresentation](http://msdn.microsoft.com/library/55ff4906-09e5-2c5c-0ed7-5f7a767542f7%28Office.15%29.aspx)property to return the presentation in the active window. The following example saves the active presentation.




```
ActivePresentation.Save
```

Use the [Presentation](http://msdn.microsoft.com/library/f009e2c3-aa08-09f0-c879-a25b8d1e0405%28Office.15%29.aspx)property to return the presentation that's in the specified document window or slide show window. The following example displays the name of the slide show running in slide show window one.




```
MsgBox SlideShowWindows(1).Presentation.Name
```


## Methods



|**Name**|
|:-----|
|[AcceptAll](http://msdn.microsoft.com/library/8212b39f-7ab1-0f30-40e7-51470574ecbe%28Office.15%29.aspx)|
|[AddTitleMaster](http://msdn.microsoft.com/library/b49baa5b-217a-ab6d-3cb3-ff74e533ef20%28Office.15%29.aspx)|
|[AddToFavorites](http://msdn.microsoft.com/library/5bdef3c1-fef2-a90b-d2be-f244e3ff1a64%28Office.15%29.aspx)|
|[ApplyTemplate](http://msdn.microsoft.com/library/0340ab20-ae21-996b-63c2-4c0b922dec6e%28Office.15%29.aspx)|
|[ApplyTemplate2](http://msdn.microsoft.com/library/43d6d14a-078f-eefa-8ad5-981b0cb6ccb9%28Office.15%29.aspx)|
|[ApplyTheme](http://msdn.microsoft.com/library/e403614b-fc39-98e0-e707-501394aacfa1%28Office.15%29.aspx)|
|[CanCheckIn](http://msdn.microsoft.com/library/26d76ca4-4fd3-2037-e193-0d2d39f59361%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/63621199-7cda-c464-527f-f55130753f08%28Office.15%29.aspx)|
|[CheckInWithVersion](http://msdn.microsoft.com/library/fc40dda4-e8cb-196d-8b82-4c0adbdf6435%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/0227528a-4693-dd1a-bb5c-cd31384014b0%28Office.15%29.aspx)|
|[Convert2](http://msdn.microsoft.com/library/001e2e98-bbdb-05cf-da93-0a9738081f08%28Office.15%29.aspx)|
|[CreateVideo](http://msdn.microsoft.com/library/d302f251-66ee-c82d-d9b9-2c29b93f7615%28Office.15%29.aspx)|
|[EndReview](http://msdn.microsoft.com/library/c77350c1-7bb5-c122-4ed2-2b2f504b517d%28Office.15%29.aspx)|
|[EnsureAllMediaUpgraded](http://msdn.microsoft.com/library/3496f149-cfd2-87b3-d69b-f7a7903bbe10%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/e114d86d-0400-35d3-fc89-d93748993874%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/bad3c9cb-49d7-2fdd-5110-9c1ed6491b08%28Office.15%29.aspx)|
|[ExportAsFixedFormat2](http://msdn.microsoft.com/library/b1101e58-e6a8-9dd4-7071-1325ba71edb1%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/411863be-0bd9-c939-1309-9f537b47f30b%28Office.15%29.aspx)|
|[GetWorkflowTasks](http://msdn.microsoft.com/library/d589e00c-3f1b-77e6-d021-b67b4d045c9a%28Office.15%29.aspx)|
|[GetWorkflowTemplates](http://msdn.microsoft.com/library/f39f2ca8-3ddc-7f45-9dea-c9c191e7cec5%28Office.15%29.aspx)|
|[LockServerFile](http://msdn.microsoft.com/library/4d32b87c-d461-392b-f267-cd2643f65fcb%28Office.15%29.aspx)|
|[Merge](http://msdn.microsoft.com/library/5cc604de-6d57-69dc-e3bc-88505b947f72%28Office.15%29.aspx)|
|[MergeWithBaseline](http://msdn.microsoft.com/library/13d9c680-fedc-7c69-5630-b814e6a7463e%28Office.15%29.aspx)|
|[NewWindow](http://msdn.microsoft.com/library/2c4e4d63-ccef-ae98-0676-fa231dec1e8c%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/57685390-43c1-4bd4-d2ee-ba34641e34c5%28Office.15%29.aspx)|
|[PublishSlides](http://msdn.microsoft.com/library/2f5c569a-fc4d-01ae-eae7-f1894541e08e%28Office.15%29.aspx)|
|[RejectAll](http://msdn.microsoft.com/library/b3f307f0-9426-d3a6-0f38-4f39ec1f6c78%28Office.15%29.aspx)|
|[RemoveDocumentInformation](http://msdn.microsoft.com/library/2c9d5cc5-8fc9-d650-b1cf-9fa3e409be1c%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/6d1251bb-27f3-0a80-bc2f-d385e2b3e3ec%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/d70a678b-66ed-9dd6-5a5e-454cdf808784%28Office.15%29.aspx)|
|[SaveCopyAs](http://msdn.microsoft.com/library/456415d1-845a-9e9b-45ce-98985e94aee5%28Office.15%29.aspx)|
|[SendFaxOverInternet](http://msdn.microsoft.com/library/4470cafb-16f5-045b-1dab-8f8ead50ffe0%28Office.15%29.aspx)|
|[SetPasswordEncryptionOptions](http://msdn.microsoft.com/library/03c07952-784b-eba6-af71-57d3d1414f81%28Office.15%29.aspx)|
|[UpdateLinks](http://msdn.microsoft.com/library/1ce2246c-d64e-c78c-8d2a-7c564eb07ecc%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/6427124b-ed76-676f-b1e9-113e82f20754%28Office.15%29.aspx)|
|[Broadcast](http://msdn.microsoft.com/library/53f0fd11-423a-cd3e-8a8f-314501acd727%28Office.15%29.aspx)|
|[BuiltInDocumentProperties](http://msdn.microsoft.com/library/d59341c4-70f4-b9be-0db6-3673d588a6bd%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/8d4b19b5-ed68-8dd4-bed3-68496230ca02%28Office.15%29.aspx)|
|[Coauthoring](http://msdn.microsoft.com/library/789469b8-d813-8038-c3e3-f8014693df79%28Office.15%29.aspx)|
|[ColorSchemes](http://msdn.microsoft.com/library/4782ee52-3bdd-4459-56da-609a92816692%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/fa8f1bb8-bac5-4579-5327-3e122d88a929%28Office.15%29.aspx)|
|[Container](http://msdn.microsoft.com/library/cc0108b7-ce95-3a1b-a400-c49700a2362c%28Office.15%29.aspx)|
|[ContentTypeProperties](http://msdn.microsoft.com/library/66bc557e-f9ca-16ca-2830-3ce5eef9a9ad%28Office.15%29.aspx)|
|[CreateVideoStatus](http://msdn.microsoft.com/library/0d4d99a9-321e-a9b7-0c58-369b66d855c3%28Office.15%29.aspx)|
|[CustomDocumentProperties](http://msdn.microsoft.com/library/3f972f15-f606-0a11-56b6-1994e617def2%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/a6bfecb1-05f8-c3f5-1356-1dd0727ab56c%28Office.15%29.aspx)|
|[CustomXMLParts](http://msdn.microsoft.com/library/72dba684-9fc2-09b3-54bb-e01c01c093c0%28Office.15%29.aspx)|
|[DefaultLanguageID](http://msdn.microsoft.com/library/8568c96c-b997-6a92-e93b-0f3d091383e2%28Office.15%29.aspx)|
|[DefaultShape](http://msdn.microsoft.com/library/318ec04a-8b30-29b3-c8a6-732564efd7a8%28Office.15%29.aspx)|
|[Designs](http://msdn.microsoft.com/library/5ad47ac9-aaab-3971-1102-fa48e8bcef8b%28Office.15%29.aspx)|
|[DisplayComments](http://msdn.microsoft.com/library/b241151a-82b5-7188-a8b8-a4a04fc37165%28Office.15%29.aspx)|
|[DocumentInspectors](http://msdn.microsoft.com/library/3f5c9fb1-de9c-170b-dca5-22215cad1dd5%28Office.15%29.aspx)|
|[DocumentLibraryVersions](http://msdn.microsoft.com/library/4c1b2055-cbbb-732d-26bd-8e6b85c26cc1%28Office.15%29.aspx)|
|[EncryptionProvider](http://msdn.microsoft.com/library/9b316f21-eeaf-4704-636f-ea68c7a36cfd%28Office.15%29.aspx)|
|[EnvelopeVisible](http://msdn.microsoft.com/library/e2a58d05-df9b-0fc6-a1d4-3349b7efa111%28Office.15%29.aspx)|
|[ExtraColors](http://msdn.microsoft.com/library/c6a9d155-206c-36e6-c180-aaff8bd85a99%28Office.15%29.aspx)|
|[FarEastLineBreakLanguage](http://msdn.microsoft.com/library/e0acc33d-0cb0-5422-4238-26b4071fb48c%28Office.15%29.aspx)|
|[FarEastLineBreakLevel](http://msdn.microsoft.com/library/fc8354a6-cbd4-d0b4-0b39-a3150afab714%28Office.15%29.aspx)|
|[Final](http://msdn.microsoft.com/library/03b16954-2f23-905b-8392-d88070e86e9f%28Office.15%29.aspx)|
|[Fonts](http://msdn.microsoft.com/library/3caece78-6ca9-bca8-5683-4722e1f563cf%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/cf6c5687-5dd0-3e71-3aa9-a370534c4117%28Office.15%29.aspx)|
|[GridDistance](http://msdn.microsoft.com/library/5c4accfe-2467-3d0e-f7f8-3e3c16d8d0ce%28Office.15%29.aspx)|
|[Guides](http://msdn.microsoft.com/library/c04540d1-57e3-e062-2518-4be6628e0166%28Office.15%29.aspx)|
|[HandoutMaster](http://msdn.microsoft.com/library/d80a8e51-61db-8da0-1fda-20a043e62569%28Office.15%29.aspx)|
|[HasHandoutMaster](http://msdn.microsoft.com/library/40834cb4-1c7a-f2f3-0027-d93f294cfec2%28Office.15%29.aspx)|
|[HasNotesMaster](http://msdn.microsoft.com/library/9dab3bbb-21c0-774a-101d-24d820b712fd%28Office.15%29.aspx)|
|[HasTitleMaster](http://msdn.microsoft.com/library/93b5932c-c03f-451a-c7f9-30683c01bcfa%28Office.15%29.aspx)|
|[HasVBProject](http://msdn.microsoft.com/library/fb8695e9-13e3-6b2e-a268-e2430e30365f%28Office.15%29.aspx)|
|[InMergeMode](http://msdn.microsoft.com/library/d9a4f840-eac2-0115-5bcf-df260b6db3c7%28Office.15%29.aspx)|
|[LayoutDirection](http://msdn.microsoft.com/library/180e6c85-618f-47e4-b0e7-f9ee3f331c25%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/a93a6d21-e3e7-0d7d-ae73-34f9511445de%28Office.15%29.aspx)|
|[NoLineBreakAfter](http://msdn.microsoft.com/library/bc9c7fd9-4aa6-b350-4c30-586a237d904a%28Office.15%29.aspx)|
|[NoLineBreakBefore](http://msdn.microsoft.com/library/d7f7f559-cf20-ef3f-60aa-122dc28da203%28Office.15%29.aspx)|
|[NotesMaster](http://msdn.microsoft.com/library/0889b69b-4c51-82cf-ccc2-ccb211d8a34e%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/81327801-ad21-967c-9682-54a847f79e29%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/0560e735-f21a-6ed3-55c6-06e025032fcb%28Office.15%29.aspx)|
|[Password](http://msdn.microsoft.com/library/977876b7-b40f-de45-c259-e91744915085%28Office.15%29.aspx)|
|[PasswordEncryptionAlgorithm](http://msdn.microsoft.com/library/728934cf-b4f3-6acd-0e42-6fc5928af807%28Office.15%29.aspx)|
|[PasswordEncryptionFileProperties](http://msdn.microsoft.com/library/086ef0bb-5307-1445-3209-f3f79927965c%28Office.15%29.aspx)|
|[PasswordEncryptionKeyLength](http://msdn.microsoft.com/library/4a3d59e4-fd4d-cd8d-8d51-cca6ebd4b758%28Office.15%29.aspx)|
|[PasswordEncryptionProvider](http://msdn.microsoft.com/library/055d4972-a835-f3fb-24df-9f275374ea6e%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/67611b54-bc31-ec2b-e645-cb3d4195bbe9%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/3f7633a8-bdab-b08d-0cf8-8df52c35865a%28Office.15%29.aspx)|
|[PrintOptions](http://msdn.microsoft.com/library/3620e0bb-1dcc-9979-d815-c3f34205aaaf%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/d0d69c81-baa0-9b33-5ee3-d8e581508a88%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/beb422cc-23c5-5de5-ed6f-0fc71315daec%28Office.15%29.aspx)|
|[Research](http://msdn.microsoft.com/library/e2d8fef9-2b21-c006-c216-2e3aee890413%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/52798ca6-e181-cf82-d397-647404235cb9%28Office.15%29.aspx)|
|[SectionProperties](http://msdn.microsoft.com/library/4b114cc6-83ef-c86d-eecc-dc39f1837a42%28Office.15%29.aspx)|
|[ServerPolicy](http://msdn.microsoft.com/library/65e50d32-96f8-63b8-6499-388bf6c61e37%28Office.15%29.aspx)|
|[SharedWorkspace](http://msdn.microsoft.com/library/79ba29b0-e51b-2644-60d7-6a044a9a7291%28Office.15%29.aspx)|
|[Signatures](http://msdn.microsoft.com/library/978e39bb-298b-d820-63cb-2924bf0770b1%28Office.15%29.aspx)|
|[SlideMaster](http://msdn.microsoft.com/library/86b11fcd-b979-6ffe-bda7-1b9c6e807d29%28Office.15%29.aspx)|
|[Slides](http://msdn.microsoft.com/library/bf481c73-3508-a074-eb2c-a5df62e55a5c%28Office.15%29.aspx)|
|[SlideShowSettings](http://msdn.microsoft.com/library/90a5a5cb-1f78-bbb2-8e4c-eb35aae13c90%28Office.15%29.aspx)|
|[SlideShowWindow](http://msdn.microsoft.com/library/9cef9c42-7a65-bd2e-3277-0145cd2cd3b9%28Office.15%29.aspx)|
|[SnapToGrid](http://msdn.microsoft.com/library/d0155913-cca5-c2ed-b1cc-6463a573ff49%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/aebb519d-ffb8-88a8-3771-5edb6b28792c%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/3b75d7ae-ce76-0023-c11e-1f39f4319ed5%28Office.15%29.aspx)|
|[TemplateName](http://msdn.microsoft.com/library/50cea27c-8181-eb32-20ae-88ae1f7ab34c%28Office.15%29.aspx)|
|[TitleMaster](http://msdn.microsoft.com/library/d5a84b2a-fff0-dcb5-e744-466428a586b5%28Office.15%29.aspx)|
|[VBASigned](http://msdn.microsoft.com/library/eebb411d-6312-f858-275f-b0f0ee12b212%28Office.15%29.aspx)|
|[VBProject](http://msdn.microsoft.com/library/76713c8c-2263-7a5a-8133-726cc94bd73a%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/ce04c680-ef68-5014-ce78-0d48d1f3b9e6%28Office.15%29.aspx)|
|[WritePassword](http://msdn.microsoft.com/library/42381e81-c5d0-3db1-f214-6619bbc6711f%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
