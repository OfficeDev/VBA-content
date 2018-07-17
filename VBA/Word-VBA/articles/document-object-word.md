---
title: Document Object (Word)
keywords: vbawd10.chm2411
f1_keywords:
- vbawd10.chm2411
ms.prod: word
api_name:
- Word.Document
ms.assetid: 8d83487a-2345-a036-a916-971c9db5b7fb
ms.date: 06/08/2017
---


# Document Object (Word)

Represents a document. The  **Document** object is a member of the **[Documents](https://msdn.microsoft.com/en-us/vba/word-vba/articles/documents-object-word)** collection. The **Documents** collection contains all the **Document** objects that are currently open in Word.


## Remarks

Use  **Documents** (Index), where Index is the document name or index number, to return a single **Document** object. The following example closes the document named "Report.doc" without saving changes.


```
Documents("Report.doc").Close SaveChanges:=wdDoNotSaveChanges
```

The index number represents the position of the document in the  **Documents** collection. The following example activates the first document in the **Documents** collection.




```
Documents(1).Activate
```

### Using ActiveDocument

You can use the  **[ActiveDocument](https://msdn.microsoft.com/en-us/vba/word-vba/articles/application-activedocument-property-word)** property to refer to the document with the focus. The following example uses the **[Activate](https://msdn.microsoft.com/en-us/vba/word-vba/articles/document-activate-method-word)** method to activate the document named "Document 1." The example also sets the page orientation to landscape mode and then prints the document.




```
Documents("Document1").Activate 
ActiveDocument.PageSetup.Orientation = wdOrientLandscape 
ActiveDocument.PrintOut
```


## Members


### Events



|**Name**|
|:-----|
|[BuildingBlockInsert](http://msdn.microsoft.com/library/6c4b1f1f-da22-63b9-a3d9-21d7bedb4b5b%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/7758dbae-b624-d3b0-f42c-1404d40ecc78%28Office.15%29.aspx)|
|[ContentControlAfterAdd](http://msdn.microsoft.com/library/9a19d147-76bd-eb92-5844-c56b2d6eae7c%28Office.15%29.aspx)|
|[ContentControlBeforeContentUpdate](http://msdn.microsoft.com/library/297241e3-fda9-1947-8b09-9dca97930dcf%28Office.15%29.aspx)|
|[ContentControlBeforeDelete](http://msdn.microsoft.com/library/a690fb97-0de3-de0e-7e84-edaaea756e83%28Office.15%29.aspx)|
|[ContentControlBeforeStoreUpdate](http://msdn.microsoft.com/library/a73aae31-bd03-1422-dbf2-1e7943d4a08a%28Office.15%29.aspx)|
|[ContentControlOnEnter](http://msdn.microsoft.com/library/593eca61-886c-85e9-fde2-1dc20c80740b%28Office.15%29.aspx)|
|[ContentControlOnExit](http://msdn.microsoft.com/library/1c988334-2bb3-2a86-747b-0d1d46894da1%28Office.15%29.aspx)|
|[New](http://msdn.microsoft.com/library/c37f7e20-f429-e921-3d17-609d536e8baa%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/80ad090c-69bf-b50e-3171-eab5414309a2%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/cc46cfdf-ae26-9bba-7084-64349859d304%28Office.15%29.aspx)|
|[XMLAfterInsert](http://msdn.microsoft.com/library/6858fd64-e96b-308e-53eb-d79595fabac7%28Office.15%29.aspx)|
|[XMLBeforeDelete](http://msdn.microsoft.com/library/1cef9cdb-a80a-8d38-9646-e3353f6c6923%28Office.15%29.aspx)|

### Methods



|**Name**|
|:-----|
|[AcceptAllRevisions](http://msdn.microsoft.com/library/3281313c-fa16-1f68-0435-f822f7cea06d%28Office.15%29.aspx)|
|[AcceptAllRevisionsShown](http://msdn.microsoft.com/library/bd9634cf-239a-2543-3681-579d4dd2f202%28Office.15%29.aspx)|
|[Activate](http://msdn.microsoft.com/library/83cc5935-020b-470a-f7aa-7fea057ec08b%28Office.15%29.aspx)|
|[AddToFavorites](http://msdn.microsoft.com/library/e810df76-18a8-d6b8-8d72-fb6386e6ce3a%28Office.15%29.aspx)|
|[ApplyDocumentTheme](http://msdn.microsoft.com/library/fd376134-f6d4-b6da-8eae-671e7e3b05e0%28Office.15%29.aspx)|
|[ApplyQuickStyleSet2](http://msdn.microsoft.com/library/7ed6e6ac-fe0f-388e-65fa-edd711d30926%28Office.15%29.aspx)|
|[ApplyTheme](http://msdn.microsoft.com/library/a4b9180e-5128-6a19-a629-47c20837f84b%28Office.15%29.aspx)|
|[AutoFormat](http://msdn.microsoft.com/library/3b81e92b-3bb8-76dc-1b58-3c70b87db664%28Office.15%29.aspx)|
|[CanCheckin](http://msdn.microsoft.com/library/7021b14b-3e45-9850-bc59-d76c267f2934%28Office.15%29.aspx)|
|[CheckConsistency](http://msdn.microsoft.com/library/9ae5e917-0bd7-7c20-ca00-eea5a7e9dff7%28Office.15%29.aspx)|
|[CheckGrammar](http://msdn.microsoft.com/library/980ddb33-94ba-fdae-3c13-6a31fdad3e14%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/3c0e5a48-65e1-c7f7-c9f1-cabaabdcb3cb%28Office.15%29.aspx)|
|[CheckInWithVersion](http://msdn.microsoft.com/library/fc041188-438e-6fab-d096-7883074a6879%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/a61a9c8b-0dee-f6e4-cefc-daca612c99c1%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/59603a58-17ee-bc65-597b-6200e8be9fbc%28Office.15%29.aspx)|
|[ClosePrintPreview](http://msdn.microsoft.com/library/8b4beae3-1893-5dbf-4463-bbce0c63b8ee%28Office.15%29.aspx)|
|[Compare](http://msdn.microsoft.com/library/2715f719-d141-c60c-8956-64aa3a58e268%28Office.15%29.aspx)|
|[ComputeStatistics](http://msdn.microsoft.com/library/f6f3c42d-b2c0-f0a8-857f-2a8e314f44fb%28Office.15%29.aspx)|
|[Convert](http://msdn.microsoft.com/library/a4392f25-c187-55d6-d3d5-ed24866a4be7%28Office.15%29.aspx)|
|[ConvertAutoHyphens](http://msdn.microsoft.com/library/ce9ad18c-881c-71c3-21bd-13c951c8e551%28Office.15%29.aspx)|
|[ConvertNumbersToText](http://msdn.microsoft.com/library/d5fed8c5-4338-04a3-6d79-c28a6ce4b9c1%28Office.15%29.aspx)|
|[ConvertVietDoc](http://msdn.microsoft.com/library/d03f0ad4-0e40-45a7-5189-1cbfa7328b2c%28Office.15%29.aspx)|
|[CopyStylesFromTemplate](http://msdn.microsoft.com/library/f02fbce7-f5aa-d71d-9043-f151f26bc9ec%28Office.15%29.aspx)|
|[CountNumberedItems](http://msdn.microsoft.com/library/b35face4-9d35-2071-90e1-628e7eca04fc%28Office.15%29.aspx)|
|[CreateLetterContent](http://msdn.microsoft.com/library/33f47344-31d2-4099-45fc-91af2d79dc7c%28Office.15%29.aspx)|
|[DataForm](http://msdn.microsoft.com/library/138f8b31-f076-8573-510f-0295fb612226%28Office.15%29.aspx)|
|[DeleteAllComments](http://msdn.microsoft.com/library/8c0bf7fa-a4de-91e0-3e2b-bb5d8897534a%28Office.15%29.aspx)|
|[DeleteAllCommentsShown](http://msdn.microsoft.com/library/b0cdbc8e-973c-1921-a646-d2f5ef091ce9%28Office.15%29.aspx)|
|[DeleteAllEditableRanges](http://msdn.microsoft.com/library/021456eb-516c-5616-3e32-19d0b9908aef%28Office.15%29.aspx)|
|[DeleteAllInkAnnotations](http://msdn.microsoft.com/library/d8446194-f86c-cb48-00e0-82ac84f9bb88%28Office.15%29.aspx)|
|[DetectLanguage](http://msdn.microsoft.com/library/625cff5b-630e-bcaa-1094-57db5029ebd9%28Office.15%29.aspx)|
|[DowngradeDocument](http://msdn.microsoft.com/library/3f79fb57-dbce-0a12-3ecf-6a1f96992d9f%28Office.15%29.aspx)|
|[EndReview](http://msdn.microsoft.com/library/bf53cefd-98e9-7e1a-016e-abd0c16e8bcd%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/fe248ff2-0a2a-b10e-fed9-d5bfb73ff1b2%28Office.15%29.aspx)|
|[FitToPages](http://msdn.microsoft.com/library/8935d286-61b7-432e-ed79-b85708dd1a01%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/ef9a3993-a7b5-5668-e804-c9d1f4fdb7dd%28Office.15%29.aspx)|
|[FreezeLayout](http://msdn.microsoft.com/library/5a61d0f3-dc28-84d1-faa1-4cfc2b32146f%28Office.15%29.aspx)|
|[GetCrossReferenceItems](http://msdn.microsoft.com/library/380e3019-2574-f50b-d871-dcebb564b06e%28Office.15%29.aspx)|
|[GetLetterContent](http://msdn.microsoft.com/library/ab0d9fa4-b193-6a7f-641d-d6f971b37457%28Office.15%29.aspx)|
|[GetWorkflowTasks](http://msdn.microsoft.com/library/6dfd67c3-f742-7979-8058-6438b1144f1f%28Office.15%29.aspx)|
|[GetWorkflowTemplates](http://msdn.microsoft.com/library/091003dc-0a26-8665-d552-0f4354313367%28Office.15%29.aspx)|
|[GoTo](http://msdn.microsoft.com/library/b03156a8-71a3-af2a-958e-79e1307e1af3%28Office.15%29.aspx)|
|[LockServerFile](http://msdn.microsoft.com/library/0e9d6d4d-0f07-d815-207e-3a1c73f8c7e7%28Office.15%29.aspx)|
|[MakeCompatibilityDefault](http://msdn.microsoft.com/library/06c3cede-312c-aacf-3780-4d79dd7c6fc3%28Office.15%29.aspx)|
|[ManualHyphenation](http://msdn.microsoft.com/library/ffd4aace-f9e3-a7ef-9dab-5694891a68ab%28Office.15%29.aspx)|
|[Merge](http://msdn.microsoft.com/library/e7ab537d-dfd3-177b-722a-6fe693c158d8%28Office.15%29.aspx)|
|[Post](http://msdn.microsoft.com/library/1ff97561-ed82-fcf3-6615-ee7ed27814fe%28Office.15%29.aspx)|
|[PresentIt](http://msdn.microsoft.com/library/2565f8a5-d99d-0b40-aea6-2ad20f9ed07f%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/bad3cd20-39e7-11b8-4a55-244bfcb6df24%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/534e3a03-b26c-5144-f6f5-09235830ec4f%28Office.15%29.aspx)|
|[Protect](http://msdn.microsoft.com/library/727bafe9-48ea-6b2f-2262-778f66487cbd%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/7dd33ac8-38f6-925d-a511-8677ca6437e6%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/0fb5671e-c933-50e6-e1fa-fe146666ad80%28Office.15%29.aspx)|
|[RejectAllRevisions](http://msdn.microsoft.com/library/d0cf9e63-0057-c832-90b5-e4057c888528%28Office.15%29.aspx)|
|[RejectAllRevisionsShown](http://msdn.microsoft.com/library/87b46681-dbc9-e38b-e20d-5da2a9a0456f%28Office.15%29.aspx)|
|[Reload](http://msdn.microsoft.com/library/4feda9b6-dd7b-2e3c-b822-04684638e9d8%28Office.15%29.aspx)|
|[ReloadAs](http://msdn.microsoft.com/library/52cab019-3084-e488-8727-24c5fd4650ce%28Office.15%29.aspx)|
|[RemoveDocumentInformation](http://msdn.microsoft.com/library/64bd3aa6-1e7f-13c1-bcc6-a9488362d7aa%28Office.15%29.aspx)|
|[RemoveLockedStyles](http://msdn.microsoft.com/library/0c20a3c9-b4b3-e9a6-06d1-a9bf9b16dc07%28Office.15%29.aspx)|
|[RemoveNumbers](http://msdn.microsoft.com/library/2f481145-f1ef-7b80-0287-3c14a5f3d2d5%28Office.15%29.aspx)|
|[RemoveTheme](http://msdn.microsoft.com/library/d9a7726b-f113-fb48-f269-f877becf0f19%28Office.15%29.aspx)|
|[Repaginate](http://msdn.microsoft.com/library/7a45ffbc-6512-6075-69a0-54a9987c27ca%28Office.15%29.aspx)|
|[ReplyWithChanges](http://msdn.microsoft.com/library/ad476bde-0240-ab4b-b246-d5b143207fa5%28Office.15%29.aspx)|
|[ResetFormFields](http://msdn.microsoft.com/library/77354799-7ba7-a4e1-5379-c7664c8820b0%28Office.15%29.aspx)|
|[ReturnToLastReadPosition](http://msdn.microsoft.com/library/d12ddc74-4557-9d7e-c47e-36311c5a748f%28Office.15%29.aspx)|
|[RunAutoMacro](http://msdn.microsoft.com/library/8eee80a6-e347-2fbb-ec86-65d09e09c764%28Office.15%29.aspx)|
|[RunLetterWizard](http://msdn.microsoft.com/library/7da6e2b9-607a-0d3e-7d0d-762a8900a486%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/7e329abc-0530-7016-7712-687de2c780a8%28Office.15%29.aspx)|
|[SaveAs2](http://msdn.microsoft.com/library/aa491007-0e31-26f5-3a5e-477381529b6e%28Office.15%29.aspx)|
|[SaveAsQuickStyleSet](http://msdn.microsoft.com/library/710dc893-235d-0571-2f5a-d5111965c6fd%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/06694b50-6a6b-ce4c-8a38-dac43ac87ba3%28Office.15%29.aspx)|
|[SelectAllEditableRanges](http://msdn.microsoft.com/library/510cd397-4c39-f36b-ed59-524247b35f16%28Office.15%29.aspx)|
|[SelectContentControlsByTag](http://msdn.microsoft.com/library/e61d5f1a-b838-e8f6-f72d-da7df327fd27%28Office.15%29.aspx)|
|[SelectContentControlsByTitle](http://msdn.microsoft.com/library/8e5fc6a8-ac06-3dee-fb83-67328765fab4%28Office.15%29.aspx)|
|[SelectLinkedControls](http://msdn.microsoft.com/library/cae4e00c-a34f-8581-07f9-b58722ec399e%28Office.15%29.aspx)|
|[SelectNodes](http://msdn.microsoft.com/library/b913720e-0f22-c626-6003-61a8dfb87f00%28Office.15%29.aspx)|
|[SelectSingleNode](http://msdn.microsoft.com/library/85f22e41-97e3-4413-c57e-26719155dc7d%28Office.15%29.aspx)|
|[SelectUnlinkedControls](http://msdn.microsoft.com/library/6d757837-0959-6754-bfae-e840ea7de339%28Office.15%29.aspx)|
|[SendFax](http://msdn.microsoft.com/library/d7a1636b-1fc2-9bfd-e7f6-828a745c06d3%28Office.15%29.aspx)|
|[SendFaxOverInternet](http://msdn.microsoft.com/library/1e1d061e-c33a-fdf1-ae63-b9a62babc1ef%28Office.15%29.aspx)|
|[SendForReview](http://msdn.microsoft.com/library/2f2cdd5c-eeca-d03f-bd58-b5586f8f461f%28Office.15%29.aspx)|
|[SendMail](http://msdn.microsoft.com/library/7e47982f-2c8f-e76b-d790-9c4e72d5110b%28Office.15%29.aspx)|
|[SetCompatibilityMode](http://msdn.microsoft.com/library/f167a640-340e-56ed-34c0-0c3dbff8575a%28Office.15%29.aspx)|
|[SetDefaultTableStyle](http://msdn.microsoft.com/library/6e932b12-6af8-af0a-5c3b-c74cefaf0d35%28Office.15%29.aspx)|
|[SetLetterContent](http://msdn.microsoft.com/library/8c9b2f6e-34a7-41a3-761d-c1a5da141aba%28Office.15%29.aspx)|
|[SetPasswordEncryptionOptions](http://msdn.microsoft.com/library/4e7c2c0a-cac2-6fa3-f237-f02c897757a1%28Office.15%29.aspx)|
|[ToggleFormsDesign](http://msdn.microsoft.com/library/4db26f6c-8e59-33b6-34eb-708b39cbed9f%28Office.15%29.aspx)|
|[TransformDocument](http://msdn.microsoft.com/library/5829a16f-b514-479f-c227-359123611970%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/f9fd64c9-aeb9-b698-6318-beb1db653ee6%28Office.15%29.aspx)|
|[UndoClear](http://msdn.microsoft.com/library/4ff5856a-ee8d-a9c8-a0a5-1d9c0a0dc9e9%28Office.15%29.aspx)|
|[Unprotect](http://msdn.microsoft.com/library/04cc2bd3-2af6-de24-bd82-7f489aefdb48%28Office.15%29.aspx)|
|[UpdateStyles](http://msdn.microsoft.com/library/fe713979-27e1-c81c-198d-5e25564233c2%28Office.15%29.aspx)|
|[ViewCode](http://msdn.microsoft.com/library/c368fce6-2fce-b2ac-6450-72dcddeec4cd%28Office.15%29.aspx)|
|[ViewPropertyBrowser](http://msdn.microsoft.com/library/937cfe62-b05d-db34-413c-61602f58eac8%28Office.15%29.aspx)|
|[WebPagePreview](http://msdn.microsoft.com/library/9e348439-3098-fe59-e501-308ad413950e%28Office.15%29.aspx)|

### Properties



|**Name**|
|:-----|
|[ActiveTheme](http://msdn.microsoft.com/library/2a68899f-8644-c9bb-1d9d-134b132eef91%28Office.15%29.aspx)|
|[ActiveThemeDisplayName](http://msdn.microsoft.com/library/b6689499-80db-12f5-8217-2c982375448b%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/707fe9e8-16de-c4aa-a0f7-6a4570d16cdd%28Office.15%29.aspx)|
|[ActiveWritingStyle](http://msdn.microsoft.com/library/035c0872-8c0b-c95f-dd0c-893982304e0f%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/8cd9178c-637a-60e3-be60-57f88e9bfc0d%28Office.15%29.aspx)|
|[AttachedTemplate](http://msdn.microsoft.com/library/e7489e88-ec82-ff16-558b-1dd5470f83c9%28Office.15%29.aspx)|
|[AutoFormatOverride](http://msdn.microsoft.com/library/85287164-98f8-fd3a-36b7-b03008e9aac3%28Office.15%29.aspx)|
|[AutoHyphenation](http://msdn.microsoft.com/library/17e53212-3717-c8a1-7f39-464622a6cd65%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/0425d9e6-1c26-3df7-bac6-6bc314a3ca47%28Office.15%29.aspx)|
|[Bibliography](http://msdn.microsoft.com/library/9538bf99-a5f4-732b-69fe-d6706451b0fc%28Office.15%29.aspx)|
|[Bookmarks](http://msdn.microsoft.com/library/47aaace6-843c-0a2d-e584-7a8ef52f6953%28Office.15%29.aspx)|
|[Broadcast](http://msdn.microsoft.com/library/cc73b751-f850-b5d0-30b3-31b78ef3f6fe%28Office.15%29.aspx)|
|[BuiltInDocumentProperties](http://msdn.microsoft.com/library/5e9a17dd-75b3-50e5-359e-dc0d0a59c46f%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/1703bbe3-6c46-a45b-9f36-1205a0d2d47c%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/3b9bb881-4e9b-d8bc-dc57-4a4be573a5a0%28Office.15%29.aspx)|
|[ClickAndTypeParagraphStyle](http://msdn.microsoft.com/library/e53d3740-265f-b3ed-350a-24dd97d9f7ab%28Office.15%29.aspx)|
|[CoAuthoring](http://msdn.microsoft.com/library/b67ac270-c583-f141-bf86-6fc385987636%28Office.15%29.aspx)|
|[CodeName](http://msdn.microsoft.com/library/684f885d-9468-9bc9-d381-ef73286330ff%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/80b82381-691b-7995-aa3e-afdf764429d6%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/1597a002-afa4-743d-60a6-ffd398f2b599%28Office.15%29.aspx)|
|[Compatibility](http://msdn.microsoft.com/library/f41979a3-8650-1807-9cf0-d1e5fdf3a49b%28Office.15%29.aspx)|
|[CompatibilityMode](http://msdn.microsoft.com/library/5e4be325-1883-7701-53a1-4d7e20e3a989%28Office.15%29.aspx)|
|[ConsecutiveHyphensLimit](http://msdn.microsoft.com/library/73ff4693-232b-fae3-8077-f6675caede1c%28Office.15%29.aspx)|
|[Container](http://msdn.microsoft.com/library/f2a0ebbe-98dc-dfc4-5879-da2b79e75b7d%28Office.15%29.aspx)|
|[Content](http://msdn.microsoft.com/library/80578329-a648-1d4b-f83d-4b2d289813fb%28Office.15%29.aspx)|
|[ContentControls](http://msdn.microsoft.com/library/86b5af56-3ab4-2440-237e-42af398b260a%28Office.15%29.aspx)|
|[ContentTypeProperties](http://msdn.microsoft.com/library/03358167-e196-3fed-58e7-cfbd9457aa2b%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/0ed9cf75-8bae-ba10-4ba0-12a73ff84c08%28Office.15%29.aspx)|
|[CurrentRsid](http://msdn.microsoft.com/library/500a743e-6d1e-e93d-b4d2-20ac13c4651a%28Office.15%29.aspx)|
|[CustomDocumentProperties](http://msdn.microsoft.com/library/4f8ac449-b9b3-45a0-7962-df7252067e67%28Office.15%29.aspx)|
|[CustomXMLParts](http://msdn.microsoft.com/library/302bbfd0-2f82-64ba-06fe-ee329c128bf6%28Office.15%29.aspx)|
|[DefaultTableStyle](http://msdn.microsoft.com/library/b6782b12-09a6-77b0-a52d-81d4028e7c19%28Office.15%29.aspx)|
|[DefaultTabStop](http://msdn.microsoft.com/library/55c7a9e4-0a25-cd32-36b0-fc9431b1d110%28Office.15%29.aspx)|
|[DefaultTargetFrame](http://msdn.microsoft.com/library/4439bf14-34da-62b6-a290-f374eeef908a%28Office.15%29.aspx)|
|[DisableFeatures](http://msdn.microsoft.com/library/40a62de3-f74e-d604-d3fc-dfb26abeb313%28Office.15%29.aspx)|
|[DisableFeaturesIntroducedAfter](http://msdn.microsoft.com/library/5714062c-ffca-8feb-6b25-52f71568ae12%28Office.15%29.aspx)|
|[DocumentInspectors](http://msdn.microsoft.com/library/db63909c-c7e3-91f1-0ebb-0c2dd9568c2c%28Office.15%29.aspx)|
|[DocumentLibraryVersions](http://msdn.microsoft.com/library/1be5fae8-0ea1-115f-3786-6979a473448b%28Office.15%29.aspx)|
|[DocumentTheme](http://msdn.microsoft.com/library/f570f807-6b36-bed8-17b4-848142c37ce7%28Office.15%29.aspx)|
|[DoNotEmbedSystemFonts](http://msdn.microsoft.com/library/435054c0-f7e3-e206-146d-7e29cce2c71d%28Office.15%29.aspx)|
|[Email](http://msdn.microsoft.com/library/dd4f6a41-3ee6-c1bf-3a2c-e00a342e0009%28Office.15%29.aspx)|
|[EmbedLinguisticData](http://msdn.microsoft.com/library/ad76bcba-dad3-6745-8cdb-a56797054af4%28Office.15%29.aspx)|
|[EmbedTrueTypeFonts](http://msdn.microsoft.com/library/ac8fb6a1-584a-2ddb-4216-53e30473ff65%28Office.15%29.aspx)|
|[EncryptionProvider](http://msdn.microsoft.com/library/ae2536e2-0852-f00d-34fe-45dba2091bdf%28Office.15%29.aspx)|
|[Endnotes](http://msdn.microsoft.com/library/3c3e87c0-ea76-8bc4-0b2e-755bff6aa14c%28Office.15%29.aspx)|
|[EnforceStyle](http://msdn.microsoft.com/library/ce2249ca-bdb0-f2b7-e9fa-a759c4507a74%28Office.15%29.aspx)|
|[Envelope](http://msdn.microsoft.com/library/00978466-69b0-a6b8-6111-5b133dd820d5%28Office.15%29.aspx)|
|[FarEastLineBreakLanguage](http://msdn.microsoft.com/library/cf868676-b880-46e9-a1b4-9cb341c63427%28Office.15%29.aspx)|
|[FarEastLineBreakLevel](http://msdn.microsoft.com/library/11642adb-2c15-a081-ae7c-d9ebe6d5b848%28Office.15%29.aspx)|
|[Fields](http://msdn.microsoft.com/library/78707979-5d25-0168-2dba-ce88a2b26f9d%28Office.15%29.aspx)|
|[Final](http://msdn.microsoft.com/library/d7b9a436-cbb3-0a09-1047-112aa30aac90%28Office.15%29.aspx)|
|[Footnotes](http://msdn.microsoft.com/library/6257f658-69f5-4223-153b-56bc3791a99d%28Office.15%29.aspx)|
|[FormattingShowClear](http://msdn.microsoft.com/library/e6a25cc8-29be-0ba4-21ba-763676cc2f90%28Office.15%29.aspx)|
|[FormattingShowFilter](http://msdn.microsoft.com/library/41509d69-9cee-bf85-6530-c5603b9c9136%28Office.15%29.aspx)|
|[FormattingShowFont](http://msdn.microsoft.com/library/ea13daf7-6b62-ad27-bf87-21dd19e90878%28Office.15%29.aspx)|
|[FormattingShowNextLevel](http://msdn.microsoft.com/library/4b358207-480f-c9fa-fd96-98fed411065f%28Office.15%29.aspx)|
|[FormattingShowNumbering](http://msdn.microsoft.com/library/2f0d8c8c-64a0-7939-e4be-99ed58ed696f%28Office.15%29.aspx)|
|[FormattingShowParagraph](http://msdn.microsoft.com/library/b2fc92be-02f5-1ed5-aa8a-76e4ed725b49%28Office.15%29.aspx)|
|[FormattingShowUserStyleName](http://msdn.microsoft.com/library/16bdfdcd-f550-9b15-d405-20bd391aa0e5%28Office.15%29.aspx)|
|[FormFields](http://msdn.microsoft.com/library/ed97fd75-0da5-b008-26c6-ea16465fddc1%28Office.15%29.aspx)|
|[FormsDesign](http://msdn.microsoft.com/library/f5ec4968-fb3e-5cca-de0b-55c36a7ae584%28Office.15%29.aspx)|
|[Frames](http://msdn.microsoft.com/library/61b7d5dc-6ab4-d29c-6c6e-daac6a2431ed%28Office.15%29.aspx)|
|[Frameset](http://msdn.microsoft.com/library/40079f4f-be1d-c8dd-5536-ccb5f570bde9%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/795a20cb-c744-6c3c-8e7f-f7a749489819%28Office.15%29.aspx)|
|[GrammarChecked](http://msdn.microsoft.com/library/30de1405-196a-e8e0-f5af-710b217ea3fd%28Office.15%29.aspx)|
|[GrammaticalErrors](http://msdn.microsoft.com/library/24e708e3-6417-f105-43d3-9be8e450f189%28Office.15%29.aspx)|
|[GridDistanceHorizontal](http://msdn.microsoft.com/library/dabff5b7-420c-ffb7-1812-eeadbdacc864%28Office.15%29.aspx)|
|[GridDistanceVertical](http://msdn.microsoft.com/library/4b3c6f15-a379-9399-fab6-ac6ec45717fa%28Office.15%29.aspx)|
|[GridOriginFromMargin](http://msdn.microsoft.com/library/137b250a-31d6-89c7-365b-285f14ae3dac%28Office.15%29.aspx)|
|[GridOriginHorizontal](http://msdn.microsoft.com/library/e4315f83-a89c-59c1-094d-4945ae2d1ce2%28Office.15%29.aspx)|
|[GridOriginVertical](http://msdn.microsoft.com/library/6fd6a060-6f25-b7c6-f4d2-b496c4d2f4b4%28Office.15%29.aspx)|
|[GridSpaceBetweenHorizontalLines](http://msdn.microsoft.com/library/79cac143-588d-d719-c653-f24852f288b6%28Office.15%29.aspx)|
|[GridSpaceBetweenVerticalLines](http://msdn.microsoft.com/library/83658d56-6724-3e34-57bb-0b9cab537985%28Office.15%29.aspx)|
|[HasPassword](http://msdn.microsoft.com/library/4234b91c-b82c-605a-5d6c-ff18aadc3689%28Office.15%29.aspx)|
|[HasVBProject](http://msdn.microsoft.com/library/1338623e-5832-b77a-cf72-f09d7c8c80de%28Office.15%29.aspx)|
|[HTMLDivisions](http://msdn.microsoft.com/library/8e383427-0777-116c-12d8-59bcc3f819d1%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/b8db5b89-0a2a-ffe9-c353-1fa77190af75%28Office.15%29.aspx)|
|[HyphenateCaps](http://msdn.microsoft.com/library/13f421aa-7e37-4f13-9b34-7ed139421e17%28Office.15%29.aspx)|
|[HyphenationZone](http://msdn.microsoft.com/library/30ea2a99-a8f5-10f4-58f9-48533bf3ec00%28Office.15%29.aspx)|
|[Indexes](http://msdn.microsoft.com/library/47a8a5d3-3c3c-81f0-8d51-5459c5bc7f89%28Office.15%29.aspx)|
|[InlineShapes](http://msdn.microsoft.com/library/049510b5-cdb3-74e8-783a-4c8fa809b876%28Office.15%29.aspx)|
|[IsInAutosave](http://msdn.microsoft.com/library/89438dfd-3b5a-e90b-5059-a62f1e47afeb%28Office.15%29.aspx)|
|[IsMasterDocument](http://msdn.microsoft.com/library/fadf30e4-9a35-40ef-0b89-ebd981577624%28Office.15%29.aspx)|
|[IsSubdocument](http://msdn.microsoft.com/library/2b7bcae0-4934-7563-34e2-d5c5ee6deaeb%28Office.15%29.aspx)|
|[JustificationMode](http://msdn.microsoft.com/library/17d1a45f-eab7-b9f4-99d7-b5a12c7acc10%28Office.15%29.aspx)|
|[KerningByAlgorithm](http://msdn.microsoft.com/library/b49416b2-bdb7-2e13-8243-9eb24cc51a2f%28Office.15%29.aspx)|
|[Kind](http://msdn.microsoft.com/library/2a2ca204-ae61-4de2-feaa-678f564b2ca0%28Office.15%29.aspx)|
|[LanguageDetected](http://msdn.microsoft.com/library/18eba980-a599-e6f0-7d73-bee6da0474be%28Office.15%29.aspx)|
|[ListParagraphs](http://msdn.microsoft.com/library/6e34e592-e745-95cd-8ffc-cd25f75db956%28Office.15%29.aspx)|
|[Lists](http://msdn.microsoft.com/library/06d5539e-f0a2-0c93-4ade-26403eb6433e%28Office.15%29.aspx)|
|[ListTemplates](http://msdn.microsoft.com/library/dc27553a-7083-4f14-ffd6-0f440982a79c%28Office.15%29.aspx)|
|[LockQuickStyleSet](http://msdn.microsoft.com/library/df5d9ecf-8aee-78d7-f64d-fb7cf0959563%28Office.15%29.aspx)|
|[LockTheme](http://msdn.microsoft.com/library/7027bf16-3398-e232-8e61-bf4a0c10806e%28Office.15%29.aspx)|
|[MailEnvelope](http://msdn.microsoft.com/library/f37a52f5-ebfe-a9b9-056e-50f6adf4c1b4%28Office.15%29.aspx)|
|[MailMerge](http://msdn.microsoft.com/library/71c144ab-b1fb-c031-2e8d-54e9802fab5d%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/5f5f8938-4dab-19fa-f339-83099c442ec4%28Office.15%29.aspx)|
|[NoLineBreakAfter](http://msdn.microsoft.com/library/287a9e9e-355e-3faf-d7fb-ee68bb0e6568%28Office.15%29.aspx)|
|[NoLineBreakBefore](http://msdn.microsoft.com/library/03d4bb24-1941-5f12-f9e5-bccdda37fb33%28Office.15%29.aspx)|
|[OMathBreakBin](http://msdn.microsoft.com/library/7ec16236-3597-232b-f640-2a9c5713865e%28Office.15%29.aspx)|
|[OMathBreakSub](http://msdn.microsoft.com/library/a361f255-1392-eddc-7771-98e9db7c291a%28Office.15%29.aspx)|
|[OMathFontName](http://msdn.microsoft.com/library/3a1c93fd-20d7-1eb9-96d5-3d13ccdde735%28Office.15%29.aspx)|
|[OMathIntSubSupLim](http://msdn.microsoft.com/library/8c27cc79-b271-112f-8281-27f0b8e3e3ae%28Office.15%29.aspx)|
|[OMathJc](http://msdn.microsoft.com/library/5ad290b1-4787-1390-d2fa-0b2e0fc0eabc%28Office.15%29.aspx)|
|[OMathLeftMargin](http://msdn.microsoft.com/library/492af100-fe93-3b9c-92fd-71425ca8e46d%28Office.15%29.aspx)|
|[OMathNarySupSubLim](http://msdn.microsoft.com/library/2d53f3e3-a5c1-f10c-2602-7b81987af7ec%28Office.15%29.aspx)|
|[OMathRightMargin](http://msdn.microsoft.com/library/2deedb5c-e1c6-d424-3a85-c95462f43b3a%28Office.15%29.aspx)|
|[OMaths](http://msdn.microsoft.com/library/bd0305db-f102-6664-3395-287495323e6d%28Office.15%29.aspx)|
|[OMathSmallFrac](http://msdn.microsoft.com/library/a34c5e4c-5804-2cac-7b75-5e163394be75%28Office.15%29.aspx)|
|[OMathWrap](http://msdn.microsoft.com/library/486fad54-d0c2-3bab-83a0-b683b2e5fbbb%28Office.15%29.aspx)|
|[OpenEncoding](http://msdn.microsoft.com/library/a147f531-de42-47c5-1a74-12ea65e64b8b%28Office.15%29.aspx)|
|[OptimizeForWord97](http://msdn.microsoft.com/library/9db75633-508c-eddb-1ee9-5c8a2e9969b2%28Office.15%29.aspx)|
|[OriginalDocumentTitle](http://msdn.microsoft.com/library/75f716ea-f944-54da-c3d9-4376c082e6f0%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/ddc90b56-f18b-3a30-23d3-24f95d9af8a6%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/ad60de6b-6287-8ea0-142e-8795f623aa29%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/f52dc8fc-fd4d-a476-da69-d57f8ad5b9fd%28Office.15%29.aspx)|
|[Password](http://msdn.microsoft.com/library/243f1735-5367-4ac9-5643-624ccf501abe%28Office.15%29.aspx)|
|[PasswordEncryptionAlgorithm](http://msdn.microsoft.com/library/5317832f-936b-5c3b-5acc-6c067563acd6%28Office.15%29.aspx)|
|[PasswordEncryptionFileProperties](http://msdn.microsoft.com/library/8da8be02-636b-bcfb-e12c-14eadf72b3f1%28Office.15%29.aspx)|
|[PasswordEncryptionKeyLength](http://msdn.microsoft.com/library/3144a2e8-f787-e38e-4322-66c6e6ac7523%28Office.15%29.aspx)|
|[PasswordEncryptionProvider](http://msdn.microsoft.com/library/473e7599-4c04-4a29-6d5c-70228900dedf%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/809b41fb-c410-5bcb-c808-780ad5232e6f%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/17a100a0-3dc4-b15d-fcb6-e7bc57d08fc6%28Office.15%29.aspx)|
|[PrintFormsData](http://msdn.microsoft.com/library/d4582018-b119-a7a3-27c4-cf4f35d00c19%28Office.15%29.aspx)|
|[PrintPostScriptOverText](http://msdn.microsoft.com/library/614e3776-c3e7-a4ca-3148-2f285229ecb2%28Office.15%29.aspx)|
|[PrintRevisions](http://msdn.microsoft.com/library/2dd7e497-70de-6bd5-7692-5757811fdec7%28Office.15%29.aspx)|
|[ProtectionType](http://msdn.microsoft.com/library/b11de5a8-8755-293e-88d4-86ce199cb57f%28Office.15%29.aspx)|
|[ReadabilityStatistics](http://msdn.microsoft.com/library/e9da9d92-bc1f-d575-07b1-3eae2749a9e5%28Office.15%29.aspx)|
|[ReadingLayoutSizeX](http://msdn.microsoft.com/library/1b77f914-ca27-8ebf-7794-3ce49f2e117b%28Office.15%29.aspx)|
|[ReadingLayoutSizeY](http://msdn.microsoft.com/library/dc2f437c-56cd-9bd6-5808-4489e48e5b90%28Office.15%29.aspx)|
|[ReadingModeLayoutFrozen](http://msdn.microsoft.com/library/5ca8aef3-82dd-81c6-9620-57f304bcbb64%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/57421a93-808f-f216-5110-0c3b80cf6e04%28Office.15%29.aspx)|
|[ReadOnlyRecommended](http://msdn.microsoft.com/library/d7190307-c58a-fa7a-7bb0-56478eac8160%28Office.15%29.aspx)|
|[RemoveDateAndTime](http://msdn.microsoft.com/library/43520dad-0374-06c9-184e-da71de304360%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/cea369d5-6ccd-8326-abdc-c834c5b17975%28Office.15%29.aspx)|
|[Research](http://msdn.microsoft.com/library/8d37d02a-c418-a2a2-1478-362ed01d76d6%28Office.15%29.aspx)|
|[RevisedDocumentTitle](http://msdn.microsoft.com/library/9783dc13-6cf5-90ac-86c4-ea4a8fc85504%28Office.15%29.aspx)|
|[Revisions](http://msdn.microsoft.com/library/26211417-b9c5-128e-1b00-cb312dd3724f%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/45bfc77d-2f8e-078c-57c1-ed3ae9f15932%28Office.15%29.aspx)|
|[SaveEncoding](http://msdn.microsoft.com/library/9a69851e-af52-d257-d887-c90bd98eeac0%28Office.15%29.aspx)|
|[SaveFormat](http://msdn.microsoft.com/library/f8d31365-1935-307f-3663-d6e769944489%28Office.15%29.aspx)|
|[SaveFormsData](http://msdn.microsoft.com/library/0f8a14be-49e9-06d4-d601-aa724c4c3c42%28Office.15%29.aspx)|
|[SaveSubsetFonts](http://msdn.microsoft.com/library/01210b29-f346-e513-6876-3dab30b940e1%28Office.15%29.aspx)|
|[Scripts](http://msdn.microsoft.com/library/5602a262-f4e2-bc9c-1457-68536adf7ac4%28Office.15%29.aspx)|
|[Sections](http://msdn.microsoft.com/library/83c3ec94-b0ef-e8a5-b17a-ad657e7197b2%28Office.15%29.aspx)|
|[Sentences](http://msdn.microsoft.com/library/41906136-815c-4dfc-ad92-c16ad420ab91%28Office.15%29.aspx)|
|[ServerPolicy](http://msdn.microsoft.com/library/dd3d41c3-588e-3a9b-049a-9f7e18402a95%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/638ab04b-2e82-afe9-3817-740f464542cc%28Office.15%29.aspx)|
|[ShowGrammaticalErrors](http://msdn.microsoft.com/library/b219a212-232c-0edb-d702-88ed4e097940%28Office.15%29.aspx)|
|[ShowSpellingErrors](http://msdn.microsoft.com/library/75b24653-f694-a5d7-bbb7-3f75f52d9e60%28Office.15%29.aspx)|
|[Signatures](http://msdn.microsoft.com/library/2f6cf537-6f7a-9cca-1d2c-39bb581630ad%28Office.15%29.aspx)|
|[SmartDocument](http://msdn.microsoft.com/library/f9671c26-208e-1682-c792-661b701308a7%28Office.15%29.aspx)|
|[SnapToGrid](http://msdn.microsoft.com/library/7aa03a0d-65f2-725b-37fe-8a421fb1e9f7%28Office.15%29.aspx)|
|[SnapToShapes](http://msdn.microsoft.com/library/b74e7a58-deee-aed2-8956-3911dd54d9ba%28Office.15%29.aspx)|
|[SpellingChecked](http://msdn.microsoft.com/library/053f8fbd-30cd-038f-e36f-d55fdd26fe13%28Office.15%29.aspx)|
|[SpellingErrors](http://msdn.microsoft.com/library/c8a987a1-3705-ea0a-103a-99b2f17f5c6b%28Office.15%29.aspx)|
|[StoryRanges](http://msdn.microsoft.com/library/6afc9e1a-950c-e1b0-15d5-73afeb72fc59%28Office.15%29.aspx)|
|[Styles](http://msdn.microsoft.com/library/30784574-92d1-a2fa-1032-6e1f8bb79ccf%28Office.15%29.aspx)|
|[StyleSheets](http://msdn.microsoft.com/library/119a2ecb-9cbd-531e-2145-fc28da798a05%28Office.15%29.aspx)|
|[StyleSortMethod](http://msdn.microsoft.com/library/188e1f2c-e5f4-3253-4051-d78cd4668f4a%28Office.15%29.aspx)|
|[Subdocuments](http://msdn.microsoft.com/library/4d0047da-03ef-67da-61ed-8bdbeaa55024%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/c48b0b07-84c6-0097-509c-ee6fb9b3784e%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/a0e09aff-af98-5d10-ba49-01ba6fcfa2d1%28Office.15%29.aspx)|
|[TablesOfAuthorities](http://msdn.microsoft.com/library/c49d1fc5-1d0a-3b6e-ab9e-62b968766cd3%28Office.15%29.aspx)|
|[TablesOfAuthoritiesCategories](http://msdn.microsoft.com/library/c7daaf7a-6002-8377-ff68-18335f441baf%28Office.15%29.aspx)|
|[TablesOfContents](http://msdn.microsoft.com/library/8c9e923d-c363-281f-d287-3501b980804e%28Office.15%29.aspx)|
|[TablesOfFigures](http://msdn.microsoft.com/library/1c386611-82f9-0a0d-71ce-dfe006d8eab5%28Office.15%29.aspx)|
|[TextEncoding](http://msdn.microsoft.com/library/a11b45c1-1829-0df0-3403-e92268d9ec81%28Office.15%29.aspx)|
|[TextLineEnding](http://msdn.microsoft.com/library/6e1f2243-473c-0294-623e-c09588645ee3%28Office.15%29.aspx)|
|[TrackFormatting](http://msdn.microsoft.com/library/b3c39567-5aed-016b-2d43-d72be55c6ebd%28Office.15%29.aspx)|
|[TrackMoves](http://msdn.microsoft.com/library/6c94cd58-dd47-313c-c04f-f04fe6f86f02%28Office.15%29.aspx)|
|[TrackRevisions](http://msdn.microsoft.com/library/c6ff8462-805d-2494-cebb-ace6fe536f40%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/8fcf6280-5fbc-10bf-95ef-7461c02102d2%28Office.15%29.aspx)|
|[UpdateStylesOnOpen](http://msdn.microsoft.com/library/7b126a45-2347-8140-25b8-861672dcc8b5%28Office.15%29.aspx)|
|[UseMathDefaults](http://msdn.microsoft.com/library/ce96cb76-0b61-32ed-4846-7a776c318639%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/34ab71eb-397e-4c14-dfbe-d3f29f84c753%28Office.15%29.aspx)|
|[Variables](http://msdn.microsoft.com/library/93af7b84-f172-6ebd-2147-e7ebc92449c5%28Office.15%29.aspx)|
|[VBASigned](http://msdn.microsoft.com/library/aa00c1ad-8c1e-5f47-de42-72db8292d5c0%28Office.15%29.aspx)|
|[VBProject](http://msdn.microsoft.com/library/bf9d4c60-8e7a-b076-b20c-0021e9352273%28Office.15%29.aspx)|
|[WebOptions](http://msdn.microsoft.com/library/038eef42-8c57-8910-d8c1-7b9937f180c5%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/bb075fd7-2dae-18c9-f49a-0c478d840b76%28Office.15%29.aspx)|
|[WordOpenXML](http://msdn.microsoft.com/library/695afe9b-843a-ef02-be21-4d733435f1df%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/0507992a-882a-81ed-c95f-5c7e26c70ebf%28Office.15%29.aspx)|
|[WritePassword](http://msdn.microsoft.com/library/e3353e68-1196-d896-d978-2c49ceca2940%28Office.15%29.aspx)|
|[WriteReserved](http://msdn.microsoft.com/library/be5d8696-9e72-f8a3-2b47-a2fde13359f9%28Office.15%29.aspx)|
|[XMLSaveThroughXSLT](http://msdn.microsoft.com/library/cc25a073-99c5-f31b-0cad-b6e4c9a7ff0c%28Office.15%29.aspx)|
|[XMLSchemaReferences](http://msdn.microsoft.com/library/7008fb35-017d-2f14-0627-9b524138137c%28Office.15%29.aspx)|
|[XMLShowAdvancedErrors](http://msdn.microsoft.com/library/56ddb6ee-f2fd-fa8e-5f07-a5af4d749652%28Office.15%29.aspx)|
|[XMLUseXSLTWhenSaving](http://msdn.microsoft.com/library/b2161a4f-9169-6927-8f37-2bc7f5a0b319%28Office.15%29.aspx)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

