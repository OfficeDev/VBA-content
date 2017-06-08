---
title: Workbook Object (Excel)
keywords: vbaxl10.chm198072
f1_keywords:
- vbaxl10.chm198072
ms.prod: excel
api_name:
- Excel.Workbook
ms.assetid: 8c00aa60-c974-eed3-0812-3c9625eb0d4c
ms.date: 06/08/2017
---


# Workbook Object (Excel)

Represents a Microsoft Excel workbook.


## Remarks

The  **Workbook** object is a member of the [Workbooks](http://msdn.microsoft.com/library/f768da57-013a-e652-0f5d-60b03aa4240a%28Office.15%29.aspx) collection. The **Workbooks** collection contains all the **Workbook** objects currently open in Microsoft Excel.


### ThisWorkbook Property

The  [ThisWorkbook](http://msdn.microsoft.com/library/04b713dd-fd93-3cbc-f10b-05b9c3d107b1%28Office.15%29.aspx) property returns the workbook where the Visual Basic code is running. In most cases, this is the same as the active workbook. However, if the Visual Basic code is part of an add-in, the **ThisWorkbook** property won't return the active workbook. In this case, the active workbook is the workbook calling the add-in, whereas the **ThisWorkbook** property returns the add-in workbook.

If you'll be creating an add-in from your Visual Basic code, you should use the  **ThisWorkbook** property to qualify any statement that must be run on the workbook you compile into the add-in.


## Example

Use  **Workbooks** ( _index_ ), where _index_ is the workbook name or index number, to return a single [Workbook](workbook-object-excel.md) object. The following example activates workbook one.


```
Workbooks(1).Activate
```

The index number denotes the order in which the workbooks were opened or created.  `Workbooks(1)` is the first workbook created, and `Workbooks(Workbooks.Count)` is the last one created. Activating a workbook doesn't change its index number. All workbooks are included in the index count, even if they're hidden.



The  **[Name](http://msdn.microsoft.com/library/55526a99-da9c-7f14-55f7-dfe9bd8ff489%28Office.15%29.aspx)** property returns the workbook name. You cannot set the name by using this property; if you need to change the name, use the **[SaveAs](http://msdn.microsoft.com/library/fbc3ce55-27a3-aa07-3fdb-77b0d611e394%28Office.15%29.aspx)** method to save the workbook under a different name. The following example activates Sheet1 in the workbook named Cogs.xls (the workbook must already be open in Microsoft Excel).




```
Workbooks("Cogs.xls").Worksheets("Sheet1").Activate
```

The  **[ActiveWorkbook](http://msdn.microsoft.com/library/637a2a30-f80c-08cd-e5c2-84716d0fff01%28Office.15%29.aspx)** property returns the workbook that's currently active. The following example sets the name of the author for the active workbook.






```
ActiveWorkbook.Author = "Jean Selva"
```

 **Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

This example emails a worksheet tab from the active workbook using a specified email address and subject. To run this code, the active worksheet must contain the email address in cell A1, the subject in cell B1, and the name of the worksheet to send in cell C1.




```
Sub SendTab()
   'Declare and initialize your variables, and turn off screen updating.
   Dim wks As Worksheet
   Application.ScreenUpdating = False
   Set wks = ActiveSheet
   
   'Copy the target worksheet, specified in cell C1, to the clipboard.
   Worksheets(Range("C1").Value).Copy
   
   'Send the content in the clipboard to the email account specified in cell A1,
   'using the subject line specified in cell B1.
   ActiveWorkbook.SendMail wks.Range("A1").Value, wks.Range("B1").Value
   
   'Do not save changes and turn screen updating back on.
   ActiveWorkbook.Close savechanges:=False
   Application.ScreenUpdating = True
End Sub
```


## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/74bb6d8c-aec8-7bb6-5c30-9a20f9a7afe8%28Office.15%29.aspx)|
|[AddinInstall](http://msdn.microsoft.com/library/671117b2-590e-9d6f-29ae-5f0bf30d4e99%28Office.15%29.aspx)|
|[AddinUninstall](http://msdn.microsoft.com/library/e35ba67b-3e04-d950-2f8b-141e478ddb67%28Office.15%29.aspx)|
|[AfterSave](http://msdn.microsoft.com/library/97fee36a-f77c-29ab-de1d-b6069b2d74d8%28Office.15%29.aspx)|
|[AfterXmlExport](http://msdn.microsoft.com/library/fe1e0a53-9f4e-ac88-58f7-fe420e57cabd%28Office.15%29.aspx)|
|[AfterXmlImport](http://msdn.microsoft.com/library/b43adf53-6b67-6127-e69d-6ea05f68b7f6%28Office.15%29.aspx)|
|[BeforeClose](http://msdn.microsoft.com/library/1c440637-8289-c6dd-24e0-1b2764fd1694%28Office.15%29.aspx)|
|[BeforePrint](http://msdn.microsoft.com/library/2c97cb32-2bb3-2848-b5ed-32d9129af080%28Office.15%29.aspx)|
|[BeforeSave](http://msdn.microsoft.com/library/dfa3e20f-1fb2-f84f-4b92-a98f22b6e637%28Office.15%29.aspx)|
|[BeforeXmlExport](http://msdn.microsoft.com/library/ee2af5de-e52f-9434-aa7c-5dc9bb102d1b%28Office.15%29.aspx)|
|[BeforeXmlImport](http://msdn.microsoft.com/library/a0a589c6-15f9-5599-c0b6-c6f881816ad6%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/6bd5411c-ac43-95cf-6755-49780ac765e9%28Office.15%29.aspx)|
|[ModelChange](http://msdn.microsoft.com/library/efe01088-273b-f9d8-ea3e-2ea1725ba7b2%28Office.15%29.aspx)|
|[NewChart](http://msdn.microsoft.com/library/76e7f325-9244-fd8c-b38d-063f0193a5e9%28Office.15%29.aspx)|
|[NewSheet](http://msdn.microsoft.com/library/5abb254d-a2c3-7dac-e79f-0de74a081ecd%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/313adc5e-0319-4ca4-cf5d-791b7184dacf%28Office.15%29.aspx)|
|[PivotTableCloseConnection](http://msdn.microsoft.com/library/e267ab5b-382e-b270-18c8-f643e03e4604%28Office.15%29.aspx)|
|[PivotTableOpenConnection](http://msdn.microsoft.com/library/b6ce12f7-7bc6-bfcc-33f4-2e8ea6e53bae%28Office.15%29.aspx)|
|[RowsetComplete](http://msdn.microsoft.com/library/05bdddba-6716-4bba-01b6-863f27623821%28Office.15%29.aspx)|
|[SheetActivate](http://msdn.microsoft.com/library/2a7c05c3-5b66-8012-5ac5-981dcfc7f947%28Office.15%29.aspx)|
|[SheetBeforeDelete](http://msdn.microsoft.com/library/42406738-0fcd-4ef7-9bd6-abcc05f5e922%28Office.15%29.aspx)|
|[SheetBeforeDoubleClick](http://msdn.microsoft.com/library/69d21025-78ef-deab-39be-b7a092d611f5%28Office.15%29.aspx)|
|[SheetBeforeRightClick](http://msdn.microsoft.com/library/d84dd9fd-85d3-009e-281b-cfc0d2874859%28Office.15%29.aspx)|
|[SheetCalculate](http://msdn.microsoft.com/library/0610bfa5-15dc-a57f-f362-cf897bd54b91%28Office.15%29.aspx)|
|[SheetChange](http://msdn.microsoft.com/library/37e727d8-255c-ac23-45d8-13a8e7639991%28Office.15%29.aspx)|
|[SheetDeactivate](http://msdn.microsoft.com/library/befde22b-69ce-c34f-2b9e-da5e026972e3%28Office.15%29.aspx)|
|[SheetFollowHyperlink](http://msdn.microsoft.com/library/be29df8c-4e8e-f719-ae1d-f91a11b89491%28Office.15%29.aspx)|
|[SheetLensGalleryRenderComplete](http://msdn.microsoft.com/library/8ac48e9f-7a15-c674-6d96-e9c1466473bc%28Office.15%29.aspx)|
|[SheetPivotTableAfterValueChange](http://msdn.microsoft.com/library/8460f5f1-d415-7aac-6a3d-fa0944036e9c%28Office.15%29.aspx)|
|[SheetPivotTableBeforeAllocateChanges](http://msdn.microsoft.com/library/2f767b5b-27fb-33de-c91d-76bbc52ea171%28Office.15%29.aspx)|
|[SheetPivotTableBeforeCommitChanges](http://msdn.microsoft.com/library/7e189a4f-a349-f862-375a-fa66311629cc%28Office.15%29.aspx)|
|[SheetPivotTableBeforeDiscardChanges](http://msdn.microsoft.com/library/e8f1ae21-c9ed-6f4d-a85c-d6768060a66f%28Office.15%29.aspx)|
|[SheetPivotTableChangeSync](http://msdn.microsoft.com/library/c280b935-3dbf-0666-b727-64d6b4ac7ebd%28Office.15%29.aspx)|
|[SheetPivotTableUpdate](http://msdn.microsoft.com/library/0b37939a-28dd-ef8b-ea5e-fc3768f8979a%28Office.15%29.aspx)|
|[SheetSelectionChange](http://msdn.microsoft.com/library/a3829af1-2917-9526-1d64-91eeb6c198ce%28Office.15%29.aspx)|
|[SheetTableUpdate](http://msdn.microsoft.com/library/609d331e-45b9-885b-a395-d80ccf4c19a5%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/ce8b77e1-a316-c0e3-f0f8-ce4ac22ec430%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/e99d955c-1975-44c3-05b3-3aa6e851083c%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/d84f0819-00df-585f-ea31-e4ab5a72950e%28Office.15%29.aspx)|
|[WindowResize](http://msdn.microsoft.com/library/6e473482-fe16-03a2-7a27-b0cd9535c3e6%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AcceptAllChanges](http://msdn.microsoft.com/library/8d8572a9-1231-c8ef-0707-72b8b5109066%28Office.15%29.aspx)|
|[Activate](http://msdn.microsoft.com/library/628e06b3-ca3f-28cb-e0fd-e696842f69f5%28Office.15%29.aspx)|
|[AddToFavorites](http://msdn.microsoft.com/library/14e1cd5a-41be-fc9a-61fa-df87698110e8%28Office.15%29.aspx)|
|[ApplyTheme](http://msdn.microsoft.com/library/11580293-22da-9154-20a0-6435b8870ac9%28Office.15%29.aspx)|
|[BreakLink](http://msdn.microsoft.com/library/1e9d70c1-908e-92eb-26b8-d6ac753cc9c2%28Office.15%29.aspx)|
|[CanCheckIn](http://msdn.microsoft.com/library/17f7cbdd-0ce0-8e3a-46f3-cb6dafaaa40a%28Office.15%29.aspx)|
|[ChangeFileAccess](http://msdn.microsoft.com/library/07f9cfc3-eece-efc1-6c03-38782ad7bcc2%28Office.15%29.aspx)|
|[ChangeLink](http://msdn.microsoft.com/library/9b2c0b82-73ff-3bdb-63df-82c0708cb703%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/f9750086-aaa6-3b04-6b51-ebcadf6b1911%28Office.15%29.aspx)|
|[CheckInWithVersion](http://msdn.microsoft.com/library/3b37cea5-8795-bcbb-9c4b-d30b2b9a095e%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/c0376cab-a2db-c606-67bf-0a4921b81e03%28Office.15%29.aspx)|
|[DeleteNumberFormat](http://msdn.microsoft.com/library/d56c2e4c-5de2-fecf-6a1f-a9fdc79943cb%28Office.15%29.aspx)|
|[EnableConnections](http://msdn.microsoft.com/library/521ebb4c-56c6-3e21-39af-4a46934790e1%28Office.15%29.aspx)|
|[EndReview](http://msdn.microsoft.com/library/cd4a445b-4731-43ba-e46a-f80f19ea5a17%28Office.15%29.aspx)|
|[ExclusiveAccess](http://msdn.microsoft.com/library/9b92ec4f-e256-7e01-6cd7-759a0d022813%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/4d72247c-bab9-3475-4792-8899c959393c%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/d070ecc9-fbb6-c146-f250-5c99b09063ec%28Office.15%29.aspx)|
|[ForwardMailer](http://msdn.microsoft.com/library/956b1746-26f2-5968-0ef7-fa3da2be974c%28Office.15%29.aspx)|
|[GetWorkflowTasks](http://msdn.microsoft.com/library/8a5ff9e0-b23a-930c-bb65-a1daa10cd946%28Office.15%29.aspx)|
|[GetWorkflowTemplates](http://msdn.microsoft.com/library/adff72bb-39ab-69ed-8a9b-defe75a5fede%28Office.15%29.aspx)|
|[HighlightChangesOptions](http://msdn.microsoft.com/library/ac69ee3e-c5ea-5ac0-418a-0b94d56a8777%28Office.15%29.aspx)|
|[LinkInfo](http://msdn.microsoft.com/library/016295a3-72c1-95b3-c259-8f286b12b73c%28Office.15%29.aspx)|
|[LinkSources](http://msdn.microsoft.com/library/6466bea0-5af8-7af0-e9d7-7595133073ae%28Office.15%29.aspx)|
|[LockServerFile](http://msdn.microsoft.com/library/be0ac600-320e-0959-bc26-5f3f4a910f5e%28Office.15%29.aspx)|
|[MergeWorkbook](http://msdn.microsoft.com/library/393790c6-3c19-7149-a999-b8712e7a6855%28Office.15%29.aspx)|
|[NewWindow](http://msdn.microsoft.com/library/ba568cee-c1cb-6e6a-8935-2cc8ce3a8400%28Office.15%29.aspx)|
|[OpenLinks](http://msdn.microsoft.com/library/cae33bab-892e-0861-e4ec-8a334097e0d1%28Office.15%29.aspx)|
|[PivotCaches](http://msdn.microsoft.com/library/0a2e7f10-c123-5c98-fb71-56868b9f8bde%28Office.15%29.aspx)|
|[Post](http://msdn.microsoft.com/library/62ecf3bc-c551-8f06-64cc-a6c141bdf172%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/3a4e7037-fcde-5a57-4a80-45f2a0994370%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/044afc4c-74d6-3ea6-1811-2c7d9cdc5b1a%28Office.15%29.aspx)|
|[Protect](http://msdn.microsoft.com/library/0e270b93-7b0b-cc68-c7c0-4002024f4292%28Office.15%29.aspx)|
|[ProtectSharing](http://msdn.microsoft.com/library/26660bc6-136a-ffc8-987e-c96db9c08231%28Office.15%29.aspx)|
|[PurgeChangeHistoryNow](http://msdn.microsoft.com/library/7ea42af1-051b-400d-cb87-0736c49d74fb%28Office.15%29.aspx)|
|[RefreshAll](http://msdn.microsoft.com/library/c1a956dc-263c-5c24-3b51-fc4af22dcd33%28Office.15%29.aspx)|
|[RejectAllChanges](http://msdn.microsoft.com/library/a53670da-370c-9ac8-75b8-008625495c2b%28Office.15%29.aspx)|
|[ReloadAs](http://msdn.microsoft.com/library/ce6a9d1a-7945-3dca-ff2d-a42289c2ccf9%28Office.15%29.aspx)|
|[RemoveDocumentInformation](http://msdn.microsoft.com/library/e668d976-108b-c627-6118-dd3384c1315c%28Office.15%29.aspx)|
|[RemoveUser](http://msdn.microsoft.com/library/f0a978a0-7bcf-3af4-a01a-831c6c854989%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/557bb3a4-c817-e942-10cf-ba252b0db498%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/c378da35-1778-44db-5c58-8d6992ca0c93%28Office.15%29.aspx)|
|[ReplyWithChanges](http://msdn.microsoft.com/library/60424d69-0062-aa5e-ea8f-4fb07086167a%28Office.15%29.aspx)|
|[ResetColors](http://msdn.microsoft.com/library/1b56a4e9-0645-fa1e-55cc-09069c6a0ff1%28Office.15%29.aspx)|
|[RunAutoMacros](http://msdn.microsoft.com/library/85dfdadf-75e6-437d-fb7a-e17681a69b35%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/466d891d-fb4c-fb0a-474b-dedb3c4ea7a7%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/fbc3ce55-27a3-aa07-3fdb-77b0d611e394%28Office.15%29.aspx)|
|[SaveAsXMLData](http://msdn.microsoft.com/library/7c4c1be3-d3a5-6e90-7750-9f371f008541%28Office.15%29.aspx)|
|[SaveCopyAs](http://msdn.microsoft.com/library/84f58488-6a2b-7fef-1472-e1b9771a60b0%28Office.15%29.aspx)|
|[SendFaxOverInternet](http://msdn.microsoft.com/library/e7d91ac4-90d2-7555-af96-dc28736da769%28Office.15%29.aspx)|
|[SendForReview](http://msdn.microsoft.com/library/3834f5b3-6d24-1bb9-27b5-052aa2e725e3%28Office.15%29.aspx)|
|[SendMail](http://msdn.microsoft.com/library/581d197c-0748-2225-2986-64aa368aab39%28Office.15%29.aspx)|
|[SendMailer](http://msdn.microsoft.com/library/e44955e1-e250-7279-19e5-e13db80ceddc%28Office.15%29.aspx)|
|[SetLinkOnData](http://msdn.microsoft.com/library/b500a579-6e4c-5712-05cf-27c6393b3bcd%28Office.15%29.aspx)|
|[SetPasswordEncryptionOptions](http://msdn.microsoft.com/library/3b6c9bfe-4cfb-1dde-fd57-07dd474df7db%28Office.15%29.aspx)|
|[ToggleFormsDesign](http://msdn.microsoft.com/library/3a6352e3-26b9-713e-ed93-a5890b37bc0a%28Office.15%29.aspx)|
|[Unprotect](http://msdn.microsoft.com/library/39387902-a8a4-7bf2-44d7-c5bde6725778%28Office.15%29.aspx)|
|[UnprotectSharing](http://msdn.microsoft.com/library/edce1744-0906-4b4e-8b98-5d1125047bff%28Office.15%29.aspx)|
|[UpdateFromFile](http://msdn.microsoft.com/library/f5148b60-9b25-8a12-5cf3-40103dcff2a3%28Office.15%29.aspx)|
|[UpdateLink](http://msdn.microsoft.com/library/2aef72cc-a820-3e91-1f46-50c739faf2bb%28Office.15%29.aspx)|
|[WebPagePreview](http://msdn.microsoft.com/library/2c88f15e-5cd3-82da-f779-55b63959a2b0%28Office.15%29.aspx)|
|[XmlImport](http://msdn.microsoft.com/library/97964c82-1fbe-7060-0a90-23c190e0b398%28Office.15%29.aspx)|
|[XmlImportXml](http://msdn.microsoft.com/library/b0edbe49-f578-ead0-8371-0196f5c515d4%28Office.15%29.aspx)|
|[CreateForecastSheet](http://msdn.microsoft.com/library/bec7b60b-7840-af15-6d5f-f5c184ea7aee%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AccuracyVersion](http://msdn.microsoft.com/library/bc81782c-662c-87ec-8381-d06e77674d0c%28Office.15%29.aspx)|
|[ActiveChart](http://msdn.microsoft.com/library/81e18252-b1fe-2487-535e-6e24c80bef24%28Office.15%29.aspx)|
|[ActiveSheet](http://msdn.microsoft.com/library/fb5578c3-64a7-edb7-4004-e608739d4c1e%28Office.15%29.aspx)|
|[ActiveSlicer](http://msdn.microsoft.com/library/d3858353-0be1-338c-e43f-1e5ffb7f37ba%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/91b30f9d-48e5-e033-8daf-416d1c0e547d%28Office.15%29.aspx)|
|[AutoUpdateFrequency](http://msdn.microsoft.com/library/dfded5c8-94d6-8a0f-24c1-248bd502850b%28Office.15%29.aspx)|
|[AutoUpdateSaveChanges](http://msdn.microsoft.com/library/06f9951d-a17a-bf88-4f6e-65835eb112f8%28Office.15%29.aspx)|
|[BuiltinDocumentProperties](http://msdn.microsoft.com/library/3efffd7d-0681-ecbc-000a-b71eceb3f92a%28Office.15%29.aspx)|
|[CalculationVersion](http://msdn.microsoft.com/library/09633164-998f-9fa7-f257-da109c369cd7%28Office.15%29.aspx)|
|[CaseSensitive](http://msdn.microsoft.com/library/6053b576-9ede-f9d8-e2bf-c012653b60a2%28Office.15%29.aspx)|
|[ChangeHistoryDuration](http://msdn.microsoft.com/library/5ebc3cc5-dffa-60cf-08cb-b2f84424c4b4%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/0aa2b1c1-0bba-f514-6158-00cdb4a5747e%28Office.15%29.aspx)|
|[Charts](http://msdn.microsoft.com/library/582d9a78-d86f-ab69-0c22-85f8a59412d9%28Office.15%29.aspx)|
|[CheckCompatibility](http://msdn.microsoft.com/library/9379c010-6756-b7ea-b4ad-5c8a4b900124%28Office.15%29.aspx)|
|[CodeName](http://msdn.microsoft.com/library/236e97b8-2bb9-c3a9-b4da-b1c327acde95%28Office.15%29.aspx)|
|[Colors](http://msdn.microsoft.com/library/60fc038b-980b-c1bc-6d1c-69d9d31a11ba%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/8d93b8cd-c4e3-b216-eda0-da4c6e573c40%28Office.15%29.aspx)|
|[ConflictResolution](http://msdn.microsoft.com/library/5142c848-0731-14d9-5913-bbaa67bf308f%28Office.15%29.aspx)|
|[Connections](http://msdn.microsoft.com/library/9c4f4ba7-dd4b-0bc2-65b7-16455014097f%28Office.15%29.aspx)|
|[ConnectionsDisabled](http://msdn.microsoft.com/library/afd53cc5-12d8-4b22-3186-1359c14f662e%28Office.15%29.aspx)|
|[Container](http://msdn.microsoft.com/library/7ad370bc-9901-3b8b-12e6-1ee57f0300e0%28Office.15%29.aspx)|
|[ContentTypeProperties](http://msdn.microsoft.com/library/a2919232-3fa2-7f37-00c2-48eb3edb10fd%28Office.15%29.aspx)|
|[CreateBackup](http://msdn.microsoft.com/library/33f05bf8-00ef-81f4-c083-30326f019cd4%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/e03bdff2-7a93-f882-31a1-1ba8dd3c1764%28Office.15%29.aspx)|
|[CustomDocumentProperties](http://msdn.microsoft.com/library/8470adbb-5b10-96ba-71f7-c667c33b6707%28Office.15%29.aspx)|
|[CustomViews](http://msdn.microsoft.com/library/286f6d5a-fb91-a339-8e74-9014ab7f4835%28Office.15%29.aspx)|
|[CustomXMLParts](http://msdn.microsoft.com/library/bd31f001-0e5d-691b-e69e-4cb91a6dbb0e%28Office.15%29.aspx)|
|[Date1904](http://msdn.microsoft.com/library/0556311c-4e45-aea3-e922-24a5830b19d4%28Office.15%29.aspx)|
|[DefaultPivotTableStyle](http://msdn.microsoft.com/library/8e2ca78a-4eb1-1b1e-c947-8a724f6d690a%28Office.15%29.aspx)|
|[DefaultSlicerStyle](http://msdn.microsoft.com/library/0f193fb8-b766-9093-9db8-8b028da108b4%28Office.15%29.aspx)|
|[DefaultTableStyle](http://msdn.microsoft.com/library/2dc86b2c-0047-53b5-3cc4-af15c36b78cf%28Office.15%29.aspx)|
|[DefaultTimelineStyle](http://msdn.microsoft.com/library/78261166-759a-8b18-c194-1f9124ca7e4e%28Office.15%29.aspx)|
|[DisplayDrawingObjects](http://msdn.microsoft.com/library/78eec8af-416d-88e6-d1f4-0b97a008f752%28Office.15%29.aspx)|
|[DisplayInkComments](http://msdn.microsoft.com/library/bce6b184-7640-f51c-1feb-1973de6ff739%28Office.15%29.aspx)|
|[DocumentInspectors](http://msdn.microsoft.com/library/26d2575f-6e61-4509-6a67-45ae576bc9fe%28Office.15%29.aspx)|
|[DocumentLibraryVersions](http://msdn.microsoft.com/library/b6338994-b5d9-ef9b-83b5-bdd47d0fd407%28Office.15%29.aspx)|
|[DoNotPromptForConvert](http://msdn.microsoft.com/library/d2af6528-4d9f-6e94-4fa6-2322098b4b17%28Office.15%29.aspx)|
|[EnableAutoRecover](http://msdn.microsoft.com/library/04a82e4d-0231-adf1-1289-35514372c995%28Office.15%29.aspx)|
|[EncryptionProvider](http://msdn.microsoft.com/library/13047af7-2e6e-6b64-05f1-8b4bd7a734ad%28Office.15%29.aspx)|
|[EnvelopeVisible](http://msdn.microsoft.com/library/d511a75a-ddd1-64f5-a09b-720657f64c09%28Office.15%29.aspx)|
|[Excel4IntlMacroSheets](http://msdn.microsoft.com/library/70a8c8d0-1169-7c3d-904e-5a32a4693f45%28Office.15%29.aspx)|
|[Excel4MacroSheets](http://msdn.microsoft.com/library/29161ab8-da75-c7b5-561d-f4423b8ab1ef%28Office.15%29.aspx)|
|[Excel8CompatibilityMode](http://msdn.microsoft.com/library/8471493b-8733-cddf-75fa-42d3d1719300%28Office.15%29.aspx)|
|[FileFormat](http://msdn.microsoft.com/library/ef722c3c-90ea-9810-b853-a3fff19d5c60%28Office.15%29.aspx)|
|[Final](http://msdn.microsoft.com/library/55d3a155-ca0c-1f7c-8612-80aac91a8eb3%28Office.15%29.aspx)|
|[ForceFullCalculation](http://msdn.microsoft.com/library/76f46d18-79e3-9828-d126-e221ae1a8157%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/83f45d15-b009-f304-ca53-4daa80c06562%28Office.15%29.aspx)|
|[FullNameURLEncoded](http://msdn.microsoft.com/library/589d98f7-e6fa-bc28-2c8f-7cb72009737a%28Office.15%29.aspx)|
|[HasPassword](http://msdn.microsoft.com/library/e3cfdc90-1e82-5556-0064-e8269ba92539%28Office.15%29.aspx)|
|[HasVBProject](http://msdn.microsoft.com/library/b4293266-40d9-a8a4-80ff-8b19ec7ed823%28Office.15%29.aspx)|
|[HighlightChangesOnScreen](http://msdn.microsoft.com/library/146f9a16-d32b-cc8f-fece-03864f0e13a2%28Office.15%29.aspx)|
|[IconSets](http://msdn.microsoft.com/library/c837d2a8-d21d-7432-a409-f49426368556%28Office.15%29.aspx)|
|[InactiveListBorderVisible](http://msdn.microsoft.com/library/a6259862-9a29-f3a5-498f-633f51ec10e6%28Office.15%29.aspx)|
|[IsAddin](http://msdn.microsoft.com/library/b8c8b9f4-4be5-0260-957e-c6450f31a0c0%28Office.15%29.aspx)|
|[IsInplace](http://msdn.microsoft.com/library/f492c09f-79d1-cde0-6cf1-db9644e41589%28Office.15%29.aspx)|
|[KeepChangeHistory](http://msdn.microsoft.com/library/3dbc322e-2b93-ae3d-cb9e-64c657fc0f80%28Office.15%29.aspx)|
|[ListChangesOnNewSheet](http://msdn.microsoft.com/library/77adf429-baa5-f2be-6139-c2b07dda5174%28Office.15%29.aspx)|
|[Mailer](http://msdn.microsoft.com/library/b020d3f6-7120-d03c-bc42-c297bcfbebf6%28Office.15%29.aspx)|
|[Model](http://msdn.microsoft.com/library/43ccdaa8-4a12-e745-88db-9db8a328ee5e%28Office.15%29.aspx)|
|[MultiUserEditing](http://msdn.microsoft.com/library/dc721463-ec34-8c52-6701-51c406beed23%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/55526a99-da9c-7f14-55f7-dfe9bd8ff489%28Office.15%29.aspx)|
|[Names](http://msdn.microsoft.com/library/26be56ec-ea12-1600-602a-eb338d4a5a8b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/4c039b5b-050f-8f4d-bc90-7982e92fb17c%28Office.15%29.aspx)|
|[Password](http://msdn.microsoft.com/library/5eaaf8cd-4344-946e-ecfa-c0f48946d2f2%28Office.15%29.aspx)|
|[PasswordEncryptionAlgorithm](http://msdn.microsoft.com/library/2745a8da-2a61-b949-115a-7f1112a0289e%28Office.15%29.aspx)|
|[PasswordEncryptionFileProperties](http://msdn.microsoft.com/library/536ad729-424e-5f81-85e9-8a6ed71fb11a%28Office.15%29.aspx)|
|[PasswordEncryptionKeyLength](http://msdn.microsoft.com/library/2662f2f5-1ad0-4a75-82c0-3268f147948a%28Office.15%29.aspx)|
|[PasswordEncryptionProvider](http://msdn.microsoft.com/library/d5bcbbf2-8de9-6725-9cac-679d6c023b34%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/f4cbf76a-2ed3-63b7-3262-45403d6f086e%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/ef04f56e-a04d-c3d9-fdda-611be7bf9d39%28Office.15%29.aspx)|
|[PersonalViewListSettings](http://msdn.microsoft.com/library/998320bf-d703-e42f-8b43-5a7b909a846d%28Office.15%29.aspx)|
|[PersonalViewPrintSettings](http://msdn.microsoft.com/library/6e4a0a9c-4eb0-d8e1-e9ce-8e9e618996b4%28Office.15%29.aspx)|
|[PivotTables](http://msdn.microsoft.com/library/b11795e0-22c8-f089-c59a-5e3d7a09d5de%28Office.15%29.aspx)|
|[PrecisionAsDisplayed](http://msdn.microsoft.com/library/4f0c8201-5b8d-5cb5-337c-944d2c7dd8d1%28Office.15%29.aspx)|
|[ProtectStructure](http://msdn.microsoft.com/library/bf721b60-0ad1-f71c-7ef4-74d2196d320e%28Office.15%29.aspx)|
|[ProtectWindows](http://msdn.microsoft.com/library/0f285fbe-2545-5c7d-9e3d-f08d57e78092%28Office.15%29.aspx)|
|[PublishObjects](http://msdn.microsoft.com/library/b6418f80-5154-6e3f-7313-222e6438c0e1%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/f3c0ec74-63af-ed76-f854-ce2382b9fcf3%28Office.15%29.aspx)|
|[ReadOnlyRecommended](http://msdn.microsoft.com/library/3cae84e4-d5f0-f01c-64d9-ec586ffdf79c%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/f5cdc655-8ba9-6dd1-ab05-028d98c11972%28Office.15%29.aspx)|
|[Research](http://msdn.microsoft.com/library/3a7ba740-314b-664b-3be6-1e8cdeded234%28Office.15%29.aspx)|
|[RevisionNumber](http://msdn.microsoft.com/library/7ea9fde5-eb89-a9b0-b637-980f1533d8ec%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/37eb8e08-2bfa-8065-2520-a71e291ab50c%28Office.15%29.aspx)|
|[SaveLinkValues](http://msdn.microsoft.com/library/ee69911f-5a4a-5c2b-c14a-cd562f3ba9f4%28Office.15%29.aspx)|
|[ServerPolicy](http://msdn.microsoft.com/library/188f6c47-35e3-bb69-cb8d-9d78b5b8fea5%28Office.15%29.aspx)|
|[ServerViewableItems](http://msdn.microsoft.com/library/2c10a647-2b2c-0640-9990-109b89444cd2%28Office.15%29.aspx)|
|[SharedWorkspace](http://msdn.microsoft.com/library/864fdee9-7149-360f-099d-e1a9b57a31db%28Office.15%29.aspx)|
|[Sheets](http://msdn.microsoft.com/library/45e4e19e-55ea-9615-231d-9435ba6d5a63%28Office.15%29.aspx)|
|[ShowConflictHistory](http://msdn.microsoft.com/library/d8588b9e-3e4b-6224-aaa7-ce0b63ff0607%28Office.15%29.aspx)|
|[ShowPivotChartActiveFields](http://msdn.microsoft.com/library/8892b134-4882-e1ff-a265-65b36af66f1a%28Office.15%29.aspx)|
|[ShowPivotTableFieldList](http://msdn.microsoft.com/library/33c74c54-27ea-d230-c640-47109bdfd4a2%28Office.15%29.aspx)|
|[Signatures](http://msdn.microsoft.com/library/b45f8036-c2d7-6113-e95c-ff78ee6a1f46%28Office.15%29.aspx)|
|[SlicerCaches](http://msdn.microsoft.com/library/1ebb7fd1-1742-815a-b4bb-4d25d6c9e705%28Office.15%29.aspx)|
|[SmartDocument](http://msdn.microsoft.com/library/19916b63-e93a-7f1e-532c-f4bbdb60622d%28Office.15%29.aspx)|
|[Styles](http://msdn.microsoft.com/library/c9a70be9-cab5-ea5f-2e3f-949b1acf43d9%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/000c9739-13ab-d6eb-c1c3-2ce721911137%28Office.15%29.aspx)|
|[TableStyles](http://msdn.microsoft.com/library/ac23db99-b2ce-0228-7808-728817736694%28Office.15%29.aspx)|
|[TemplateRemoveExtData](http://msdn.microsoft.com/library/9851df1d-4e83-525a-8a43-bd84b0a94c74%28Office.15%29.aspx)|
|[Theme](http://msdn.microsoft.com/library/1208f610-8c6f-9a62-3378-9566a7ee6b37%28Office.15%29.aspx)|
|[UpdateLinks](http://msdn.microsoft.com/library/c8d374d7-0b32-eb32-fa29-ab496d6786e7%28Office.15%29.aspx)|
|[UpdateRemoteReferences](http://msdn.microsoft.com/library/055c1a88-c189-ddd3-c9b2-9458817cec90%28Office.15%29.aspx)|
|[UserStatus](http://msdn.microsoft.com/library/0df24f8a-b60b-fd8c-3436-903652487a09%28Office.15%29.aspx)|
|[UseWholeCellCriteria](http://msdn.microsoft.com/library/b65093aa-37ca-2aa1-4f18-c90bc7536f74%28Office.15%29.aspx)|
|[UseWildcards](http://msdn.microsoft.com/library/92e7463c-6dbe-c409-461a-ca730402ad62%28Office.15%29.aspx)|
|[VBASigned](http://msdn.microsoft.com/library/6e93161c-2fa4-1064-9b5d-a8eb96ad2bea%28Office.15%29.aspx)|
|[VBProject](http://msdn.microsoft.com/library/1bef5b7e-e169-fa4b-9810-6cd87ecd0a8d%28Office.15%29.aspx)|
|[WebOptions](http://msdn.microsoft.com/library/801742a2-f5d8-5311-ea24-fd428532ba80%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/2352d6c9-720e-b58d-6e7c-049bf21a090d%28Office.15%29.aspx)|
|[Worksheets](http://msdn.microsoft.com/library/8b7d660d-ca49-0bd0-dc57-64defa47bd5e%28Office.15%29.aspx)|
|[WritePassword](http://msdn.microsoft.com/library/ac89063e-6ef5-f7c5-abb0-4e6ef1c5fd05%28Office.15%29.aspx)|
|[WriteReserved](http://msdn.microsoft.com/library/96cc86d1-0e77-b6f3-3045-f6346de0f969%28Office.15%29.aspx)|
|[WriteReservedBy](http://msdn.microsoft.com/library/f053c197-3af3-9ab7-bee1-f72ee311a5b8%28Office.15%29.aspx)|
|[XmlMaps](http://msdn.microsoft.com/library/c7893167-bfa1-e1df-58f3-782b80322fad%28Office.15%29.aspx)|
|[XmlNamespaces](http://msdn.microsoft.com/library/b93aba02-f831-6321-1c0d-2a30d417e57f%28Office.15%29.aspx)|
|[Queries](http://msdn.microsoft.com/library/29ee16cb-b6f2-2358-7e1a-3b1e7f9cf654%28Office.15%29.aspx)|

## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributor"> </a>


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
