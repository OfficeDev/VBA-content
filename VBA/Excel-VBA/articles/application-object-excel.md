---
title: Application Object (Excel)
keywords: vbaxl10.chm182073
f1_keywords:
- vbaxl10.chm182073
ms.prod: excel
api_name:
- Excel.Application
ms.assetid: 19b73597-5cf9-4f56-8227-b5211f657f6f
ms.date: 06/08/2017
---

# Application Object (Excel)

Represents the entire Microsoft Excel application.


## Example

Use the **Application** property to return the **Application** object. The following example applies the **Windows** property to the **Application** object.

```
Application.Windows("book1.xls").Activate
```

<br/>

The following example creates an Excel workbook object in another application and then opens a workbook in Excel.

```
Set xl = CreateObject("Excel.Sheet") 
xl.Application.Workbooks.Open "newbook.xls"
```

<br/>

Many of the properties and methods that return the most common user-interface objects, such as the active cell (**ActiveCell** property), can be used without the **Application** object qualifier. For example, instead of writing:

```
Application.ActiveCell.Font.Bold = True
```

You can write: 

```
ActiveCell.Font.Bold = True
```


## Remarks

The  **Application** object contains:

- Application-wide settings and options.
    
- Methods that return top-level objects, such as **[ActiveCell](http://msdn.microsoft.com/library/7ebfbec8-dc4e-36c5-188a-347d42649e76%28Office.15%29.aspx)**, **[ActiveSheet](http://msdn.microsoft.com/library/6ed42d87-2ad5-eecc-ad5b-4c92617a04bc%28Office.15%29.aspx)**, and so on.
    
## Events

|**Name**|
|:-----|
|[AfterCalculate](http://msdn.microsoft.com/library/ed76a36f-1b52-4464-da44-e64c81fb8d38%28Office.15%29.aspx)|
|[NewWorkbook](http://msdn.microsoft.com/library/a3c29269-af09-08da-f0c3-82e192aa896f%28Office.15%29.aspx)|
|[ProtectedViewWindowActivate](http://msdn.microsoft.com/library/271e0344-9dd1-bf08-f7bd-9892ca6ad450%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeClose](http://msdn.microsoft.com/library/5fa37062-61c7-3002-1ea0-c5bd396b6a9b%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeEdit](http://msdn.microsoft.com/library/b823b4a4-5d2f-7caf-f66f-5053b58082e4%28Office.15%29.aspx)|
|[ProtectedViewWindowDeactivate](http://msdn.microsoft.com/library/39df50ca-53e0-784a-a803-e9ac6f456d11%28Office.15%29.aspx)|
|[ProtectedViewWindowOpen](http://msdn.microsoft.com/library/17c847d9-a9d2-28da-832a-01d7719f1248%28Office.15%29.aspx)|
|[ProtectedViewWindowResize](http://msdn.microsoft.com/library/9ecac960-8ed3-f0be-9e93-3793c49d2b76%28Office.15%29.aspx)|
|[SheetActivate](http://msdn.microsoft.com/library/06387251-ba01-531c-56c8-359ffb0940e5%28Office.15%29.aspx)|
|[SheetBeforeDelete](http://msdn.microsoft.com/library/9544d9db-6bb0-43bb-91f3-3f0075c3e03b%28Office.15%29.aspx)|
|[SheetBeforeDoubleClick](http://msdn.microsoft.com/library/969394a3-2c87-36a5-2d64-521bad8849be%28Office.15%29.aspx)|
|[SheetBeforeRightClick](http://msdn.microsoft.com/library/eb91ede3-3f17-7cf8-2b6f-b519acd11ce3%28Office.15%29.aspx)|
|[SheetCalculate](http://msdn.microsoft.com/library/8d0c9042-2bf7-3575-dedb-4f99e1391de1%28Office.15%29.aspx)|
|[SheetChange](http://msdn.microsoft.com/library/0b06ad02-52c0-f0a3-c827-b7e51aecc81c%28Office.15%29.aspx)|
|[SheetDeactivate](http://msdn.microsoft.com/library/7596a2ab-4626-eb05-3b3d-64e6d9e142b8%28Office.15%29.aspx)|
|[SheetFollowHyperlink](http://msdn.microsoft.com/library/656e0ec6-64ea-1685-f088-a7e30bfaef38%28Office.15%29.aspx)|
|[SheetLensGalleryRenderComplete](http://msdn.microsoft.com/library/0b0c8d91-83dd-f4ee-82de-25ac739802b1%28Office.15%29.aspx)|
|[SheetPivotTableAfterValueChange](http://msdn.microsoft.com/library/07cab356-1a13-a839-7344-a4de99dba55e%28Office.15%29.aspx)|
|[SheetPivotTableBeforeAllocateChanges](http://msdn.microsoft.com/library/b76cc20d-6251-def7-44d2-504fd6e9cda9%28Office.15%29.aspx)|
|[SheetPivotTableBeforeCommitChanges](http://msdn.microsoft.com/library/ba586d2e-772a-24e3-0886-fb309f17ebf6%28Office.15%29.aspx)|
|[SheetPivotTableBeforeDiscardChanges](http://msdn.microsoft.com/library/8623adc6-d256-bebb-fe35-8710390af19f%28Office.15%29.aspx)|
|[SheetPivotTableUpdate](http://msdn.microsoft.com/library/f42d1f7b-6395-326b-4b4f-72b497c81bd3%28Office.15%29.aspx)|
|[SheetSelectionChange](http://msdn.microsoft.com/library/c98203c2-b306-d8b7-b75f-1304be7b5751%28Office.15%29.aspx)|
|[SheetTableUpdate](http://msdn.microsoft.com/library/6b8a5015-d715-0921-2292-be373670f82e%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/5c618983-27d8-49b1-0a52-001c7a1f94d8%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/6adcba54-3d4a-f780-915e-5798303faf60%28Office.15%29.aspx)|
|[WindowResize](http://msdn.microsoft.com/library/937c4b8f-3b37-ada7-ee72-0ad4707c2e2b%28Office.15%29.aspx)|
|[WorkbookActivate](http://msdn.microsoft.com/library/a2b6ea2e-3753-69bf-9a81-ec2fce29d4fd%28Office.15%29.aspx)|
|[WorkbookAddinInstall](http://msdn.microsoft.com/library/955c8f2a-4647-ed7e-29f9-8d6d165898ec%28Office.15%29.aspx)|
|[WorkbookAddinUninstall](http://msdn.microsoft.com/library/8c02eb17-e966-703d-36ed-30ce43a56275%28Office.15%29.aspx)|
|[WorkbookAfterSave](http://msdn.microsoft.com/library/4efa76bd-dd9f-3c7b-efa1-e1815ac8774d%28Office.15%29.aspx)|
|[WorkbookAfterXmlExport](http://msdn.microsoft.com/library/9d542c67-4244-d018-4db6-3584f0caec7c%28Office.15%29.aspx)|
|[WorkbookAfterXmlImport](http://msdn.microsoft.com/library/a58cc327-3776-fe5b-68d4-406269f30379%28Office.15%29.aspx)|
|[WorkbookBeforeClose](http://msdn.microsoft.com/library/9c3618ea-0e5e-e4fe-20af-279826bfa7c3%28Office.15%29.aspx)|
|[WorkbookBeforePrint](http://msdn.microsoft.com/library/27cb5f84-fda3-dc89-6e12-0c31ed16f47c%28Office.15%29.aspx)|
|[WorkbookBeforeSave](http://msdn.microsoft.com/library/e93a7cef-b018-ddab-c96f-b3215143f31f%28Office.15%29.aspx)|
|[WorkbookBeforeXmlExport](http://msdn.microsoft.com/library/2c228d28-2d42-40b0-ee36-214bc720d78a%28Office.15%29.aspx)|
|[WorkbookBeforeXmlImport](http://msdn.microsoft.com/library/33c7f386-9eec-6ba4-519e-9480ab2f5a31%28Office.15%29.aspx)|
|[WorkbookDeactivate](http://msdn.microsoft.com/library/0a6a55ea-5374-4de7-e48e-e52d903cc749%28Office.15%29.aspx)|
|[WorkbookModelChange](http://msdn.microsoft.com/library/62a32a29-e052-e812-82a7-58bdabadd80f%28Office.15%29.aspx)|
|[WorkbookNewChart](http://msdn.microsoft.com/library/8456e472-6ea5-a916-10d6-f12afefb58fc%28Office.15%29.aspx)|
|[WorkbookNewSheet](http://msdn.microsoft.com/library/5190254f-b7f4-10e5-41f5-704b1466ff68%28Office.15%29.aspx)|
|[WorkbookOpen](http://msdn.microsoft.com/library/37a5b55d-7968-29a2-3f87-edc3334c8ced%28Office.15%29.aspx)|
|[WorkbookPivotTableCloseConnection](http://msdn.microsoft.com/library/4c1d4cb2-f589-3c3c-ab4c-dcb08467fcfb%28Office.15%29.aspx)|
|[WorkbookPivotTableOpenConnection](http://msdn.microsoft.com/library/5f07e995-96fd-86ac-2d1c-1366528fd8c6%28Office.15%29.aspx)|
|[WorkbookRowsetComplete](http://msdn.microsoft.com/library/cc472400-5622-5b4f-60a2-d3347ded266f%28Office.15%29.aspx)|
|[WorkbookSync](http://msdn.microsoft.com/library/ca23985c-e5ea-d2cb-bce3-2b52c5dff3a1%28Office.15%29.aspx)|

<br/>

## Methods

|**Name**|
|:-----|
|[ActivateMicrosoftApp](http://msdn.microsoft.com/library/e11d8165-5aad-2b1d-f9d1-797038d96afb%28Office.15%29.aspx)|
|[AddCustomList](http://msdn.microsoft.com/library/31518c3c-78ce-f9e9-9572-a1338aa6d2e7%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/2818a08b-1c02-9f10-db03-db509a251f60%28Office.15%29.aspx)|
|[CalculateFull](http://msdn.microsoft.com/library/11be6386-a5de-817f-0624-b7e7fd502ed3%28Office.15%29.aspx)|
|[CalculateFullRebuild](http://msdn.microsoft.com/library/6d3dac24-7fb8-05fd-b6ee-cb3ef7d5f33a%28Office.15%29.aspx)|
|[CalculateUntilAsyncQueriesDone](http://msdn.microsoft.com/library/5796365e-5a79-3e4b-023e-3a1a120e92df%28Office.15%29.aspx)|
|[CentimetersToPoints](http://msdn.microsoft.com/library/2693973c-7d80-8883-6959-afabdb51b9b2%28Office.15%29.aspx)|
|[CheckAbort](http://msdn.microsoft.com/library/e407aeff-b401-029a-9ada-8f11eef54fb0%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/dfae0789-4635-5ec5-5146-c5a1acefa306%28Office.15%29.aspx)|
|[ConvertFormula](http://msdn.microsoft.com/library/6ed0a76c-9db5-f6ab-a91d-d4e1b6674c53%28Office.15%29.aspx)|
|[DDEExecute](http://msdn.microsoft.com/library/18cd97e6-4dff-2386-84bf-25e8c85b5277%28Office.15%29.aspx)|
|[DDEInitiate](http://msdn.microsoft.com/library/4b14e2ee-d7b0-a028-42a7-0809fa381f7e%28Office.15%29.aspx)|
|[DDEPoke](http://msdn.microsoft.com/library/5d00e0da-e041-7a9e-3b55-f5edd3f2a4a0%28Office.15%29.aspx)|
|[DDERequest](http://msdn.microsoft.com/library/822ef77e-5f11-aced-f770-05175ce128c7%28Office.15%29.aspx)|
|[DDETerminate](http://msdn.microsoft.com/library/f05adf6d-5714-12c4-39ce-af4bc31f4d32%28Office.15%29.aspx)|
|[DeleteCustomList](http://msdn.microsoft.com/library/41a936f7-05b5-520f-f5c5-172a5ea124d9%28Office.15%29.aspx)|
|[DisplayXMLSourcePane](http://msdn.microsoft.com/library/1dea98ac-8d36-4745-cb6a-9a607e863ff2%28Office.15%29.aspx)|
|[DoubleClick](http://msdn.microsoft.com/library/17958601-3e24-a7bb-7d8c-0c45b955f449%28Office.15%29.aspx)|
|[Evaluate](http://msdn.microsoft.com/library/031ce9e0-a7af-30f3-aa9f-fc776d8bf146%28Office.15%29.aspx)|
|[ExecuteExcel4Macro](http://msdn.microsoft.com/library/0afa77ab-43e0-0120-4ffd-25e290c72f6c%28Office.15%29.aspx)|
|[FindFile](http://msdn.microsoft.com/library/c4367047-0f7d-1bda-5103-f26eedd98e8a%28Office.15%29.aspx)|
|[GetCustomListContents](http://msdn.microsoft.com/library/3adafb35-f7d0-0233-ff7c-c31d5e48f574%28Office.15%29.aspx)|
|[GetCustomListNum](http://msdn.microsoft.com/library/c4a97a96-333a-1021-7324-5cca4f0d9f3c%28Office.15%29.aspx)|
|[GetOpenFilename](http://msdn.microsoft.com/library/83931dc2-59b3-550b-6ce1-880413fd23d6%28Office.15%29.aspx)|
|[GetPhonetic](http://msdn.microsoft.com/library/530be07e-04ed-81c5-3b12-93b78e494a3b%28Office.15%29.aspx)|
|[GetSaveAsFilename](http://msdn.microsoft.com/library/2ad52070-22d7-a755-9267-daaa5edbbb0d%28Office.15%29.aspx)|
|[Goto](http://msdn.microsoft.com/library/ce60e6d4-18e5-056c-229e-8c0b730109ae%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/e54291a6-96a5-cc55-72de-f2c1800391e2%28Office.15%29.aspx)|
|[InchesToPoints](http://msdn.microsoft.com/library/7689eae4-f533-32e3-d431-4873029a8bc1%28Office.15%29.aspx)|
|[InputBox](http://msdn.microsoft.com/library/d3bd2f3a-7fed-20fa-918d-a71e2a2a1d49%28Office.15%29.aspx)|
|[Intersect](http://msdn.microsoft.com/library/856d052a-3207-ced2-941c-b466cb880a93%28Office.15%29.aspx)|
|[MacroOptions](http://msdn.microsoft.com/library/c81abbc5-0865-9e86-f188-652c88ac6baa%28Office.15%29.aspx)|
|[MailLogoff](http://msdn.microsoft.com/library/5265e9c1-6c04-3591-7133-5274e5b56347%28Office.15%29.aspx)|
|[MailLogon](http://msdn.microsoft.com/library/0a6c8752-739d-b996-1426-4d3021ea5323%28Office.15%29.aspx)|
|[NextLetter](http://msdn.microsoft.com/library/002ace38-48f1-cac2-6bbb-428b119c8ed0%28Office.15%29.aspx)|
|[OnKey](http://msdn.microsoft.com/library/43662d8b-19e2-2b4a-4c3a-c64be4007643%28Office.15%29.aspx)|
|[OnRepeat](http://msdn.microsoft.com/library/7d535e14-c779-af87-60eb-68ec8e651459%28Office.15%29.aspx)|
|[OnTime](http://msdn.microsoft.com/library/31268da0-8ec7-7169-a1d0-8db34b3385cd%28Office.15%29.aspx)|
|[OnUndo](http://msdn.microsoft.com/library/12e59bbb-e134-3728-7c8d-629dcda0e908%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/d01de494-95c7-6e3e-3049-f89b31aa9d0c%28Office.15%29.aspx)|
|[RecordMacro](http://msdn.microsoft.com/library/8b6c9757-b589-04e6-5650-edfc4104e517%28Office.15%29.aspx)|
|[RegisterXLL](http://msdn.microsoft.com/library/b0d97511-bb81-7c6a-7bbb-3f87c4364e95%28Office.15%29.aspx)|
|[Repeat](http://msdn.microsoft.com/library/ce8f6340-174e-b6cf-0f99-f39be2cde5c2%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/3e0167ab-b101-018f-0f89-ada116b8bb72%28Office.15%29.aspx)|
|[SendKeys](http://msdn.microsoft.com/library/585666b9-adbc-5d04-c480-58e55ea7fb9d%28Office.15%29.aspx)|
|[SharePointVersion](http://msdn.microsoft.com/library/9d561b10-dba9-8af5-6e64-66e41714e894%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/b56bb8a0-2cd1-356a-03ba-47eb6f56f455%28Office.15%29.aspx)|
|[Union](http://msdn.microsoft.com/library/7c70a5be-2696-5fc2-bd69-6c6ff4d3291e%28Office.15%29.aspx)|
|[Volatile](http://msdn.microsoft.com/library/27047561-9d76-b37d-100d-1c58e6edf494%28Office.15%29.aspx)|
|[Wait](http://msdn.microsoft.com/library/71425d1c-6b37-a510-d8b5-072136e98f04%28Office.15%29.aspx)|

<br/>

## Properties

|**Name**|
|:-----|
|[ActiveCell](http://msdn.microsoft.com/library/7ebfbec8-dc4e-36c5-188a-347d42649e76%28Office.15%29.aspx)|
|[ActiveChart](http://msdn.microsoft.com/library/37b1901c-a9c2-4a86-ce05-22f3989bc9d8%28Office.15%29.aspx)|
|[ActiveEncryptionSession](http://msdn.microsoft.com/library/a13b5785-5b27-6276-1df5-f213a419446d%28Office.15%29.aspx)|
|[ActivePrinter](http://msdn.microsoft.com/library/72c4a525-27ab-f109-64d3-bcc7a12df505%28Office.15%29.aspx)|
|[ActiveProtectedViewWindow](http://msdn.microsoft.com/library/2202c3b4-8880-7a26-8a56-8f2d2e7b7343%28Office.15%29.aspx)|
|[ActiveSheet](http://msdn.microsoft.com/library/6ed42d87-2ad5-eecc-ad5b-4c92617a04bc%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/8f788ad0-ae4e-2442-593c-9440e37100de%28Office.15%29.aspx)|
|[ActiveWorkbook](http://msdn.microsoft.com/library/637a2a30-f80c-08cd-e5c2-84716d0fff01%28Office.15%29.aspx)|
|[AddIns](http://msdn.microsoft.com/library/0798690a-910a-b832-e143-df51d7c061ca%28Office.15%29.aspx)|
|[AddIns2](http://msdn.microsoft.com/library/3fd3de81-beae-c5b0-572d-c3f81e251db2%28Office.15%29.aspx)|
|[AlertBeforeOverwriting](http://msdn.microsoft.com/library/75c69d9d-bd6e-c0c9-71c4-c9d92333d233%28Office.15%29.aspx)|
|[AltStartupPath](http://msdn.microsoft.com/library/92c987ed-542d-c227-d9c3-de64eba325e0%28Office.15%29.aspx)|
|[AlwaysUseClearType](http://msdn.microsoft.com/library/f848fedf-8dc4-83c5-e2c6-e20db4d0eb0b%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/03452379-293c-2e36-ad97-bfd3de47147a%28Office.15%29.aspx)|
|[ArbitraryXMLSupportAvailable](http://msdn.microsoft.com/library/f63a64fa-5293-712a-bbbd-5dc07abda8da%28Office.15%29.aspx)|
|[AskToUpdateLinks](http://msdn.microsoft.com/library/1d04eb45-9dcc-e15f-daf1-a6ce9fa737ae%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/bfb1fe5e-a87d-e54c-dc2f-5dd308dc8a8b%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/e339617e-e086-7324-9240-4db9cfcfcee5%28Office.15%29.aspx)|
|[AutoFormatAsYouTypeReplaceHyperlinks](http://msdn.microsoft.com/library/92c02556-f39a-7ca4-31f5-88a5c9da12ea%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/ae19bf93-dc0f-f18a-d8ce-f54108602844%28Office.15%29.aspx)|
|[AutoPercentEntry](http://msdn.microsoft.com/library/80ade0a1-84ae-5a17-6a75-189c0c06843d%28Office.15%29.aspx)|
|[AutoRecover](http://msdn.microsoft.com/library/bc2453fa-4319-c1da-5ad5-2efb306c3063%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/da8ec8af-c1d9-917e-a057-a4762a783124%28Office.15%29.aspx)|
|[CalculateBeforeSave](http://msdn.microsoft.com/library/133dbe08-8f41-c07c-8362-48412ed7c086%28Office.15%29.aspx)|
|[Calculation](http://msdn.microsoft.com/library/5ae7f2dd-e79a-a4ee-f701-2fff1b77f499%28Office.15%29.aspx)|
|[CalculationInterruptKey](http://msdn.microsoft.com/library/1187c122-0498-a82c-5479-1595c7f06448%28Office.15%29.aspx)|
|[CalculationState](http://msdn.microsoft.com/library/2f424286-7757-12e2-77c2-c26cf7c4bcf1%28Office.15%29.aspx)|
|[CalculationVersion](http://msdn.microsoft.com/library/10de3816-9873-09e5-4141-effdbfe5cd9c%28Office.15%29.aspx)|
|[Caller](http://msdn.microsoft.com/library/0cfec08d-3cbc-0ab1-419a-f5b5702c3969%28Office.15%29.aspx)|
|[CanPlaySounds](http://msdn.microsoft.com/library/4e74bdbe-c649-9171-b42c-3c226b6c92a0%28Office.15%29.aspx)|
|[CanRecordSounds](http://msdn.microsoft.com/library/a2175b38-ee89-2e92-ffaa-c550115e319b%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/618f5623-2eb7-4b7e-2f15-c30a0c2e0fe2%28Office.15%29.aspx)|
|[CellDragAndDrop](http://msdn.microsoft.com/library/da10e4ce-905b-5cc3-75ff-e3248cdf6391%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/9788c893-13c3-eb57-bcf7-50806b476ba3%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/124b4d82-de33-c5df-7aa0-1a9c3484a680%28Office.15%29.aspx)|
|[Charts](http://msdn.microsoft.com/library/d60d912c-7c70-7004-d876-e83b98a13de9%28Office.15%29.aspx)|
|[ClipboardFormats](http://msdn.microsoft.com/library/9b0de0b9-6acf-a73c-6d29-a405d0784170%28Office.15%29.aspx)|
|[ClusterConnector](http://msdn.microsoft.com/library/5382b95a-c796-e638-5c11-5524e4be3acb%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/242d9112-9352-c3a6-e23e-59aec3d8f68f%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/d51f3373-ba5d-20b4-7557-246a6fcf89c3%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/b1884d43-557b-47be-1cef-20404069b576%28Office.15%29.aspx)|
|[CommandUnderlines](http://msdn.microsoft.com/library/07d3ea82-6ef4-db6f-f3cf-bef992664408%28Office.15%29.aspx)|
|[ConstrainNumeric](http://msdn.microsoft.com/library/910dd5ad-1750-71b8-8c12-df5107d21063%28Office.15%29.aspx)|
|[ControlCharacters](http://msdn.microsoft.com/library/039a266a-e5ae-468e-e3ee-101fa2b12863%28Office.15%29.aspx)|
|[CopyObjectsWithCells](http://msdn.microsoft.com/library/86836569-7bd1-bfe7-2def-6cf43a7c0368%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/92ceed4a-4e47-18d5-6023-f1018eefd071%28Office.15%29.aspx)|
|[Cursor](http://msdn.microsoft.com/library/5137b89d-aba9-3e5f-b6c4-cd2264a7bd7f%28Office.15%29.aspx)|
|[CursorMovement](http://msdn.microsoft.com/library/4be5a3fd-7a68-1190-5888-239497d53cb1%28Office.15%29.aspx)|
|[CustomListCount](http://msdn.microsoft.com/library/98a32161-e413-a0b7-a6be-4b11ae90fc00%28Office.15%29.aspx)|
|[CutCopyMode](http://msdn.microsoft.com/library/d45d3352-2a33-99ae-22f2-0b1c11466209%28Office.15%29.aspx)|
|[DataEntryMode](http://msdn.microsoft.com/library/1fd9f191-3aa5-2548-2d41-b9d2bc79654b%28Office.15%29.aspx)|
|[DDEAppReturnCode](http://msdn.microsoft.com/library/9b55dcce-eea8-a8b7-dace-296191de18a4%28Office.15%29.aspx)|
|[DecimalSeparator](http://msdn.microsoft.com/library/2423d0dd-2b67-e8d2-c611-2bd3c8061f66%28Office.15%29.aspx)|
|[DefaultFilePath](http://msdn.microsoft.com/library/8eb8f6a2-f5fe-0b7e-172f-e7cfabef4af2%28Office.15%29.aspx)|
|[DefaultSaveFormat](http://msdn.microsoft.com/library/bb5c50db-8ba3-f79a-4577-f293ebc52b50%28Office.15%29.aspx)|
|[DefaultSheetDirection](http://msdn.microsoft.com/library/33fad777-e2dd-99b5-9b33-a573a729b331%28Office.15%29.aspx)|
|[DefaultWebOptions](http://msdn.microsoft.com/library/51524888-0812-85ee-c8f9-e14d9b558f57%28Office.15%29.aspx)|
|[DeferAsyncQueries](http://msdn.microsoft.com/library/21f05a5a-40e8-304a-f537-41ea171a114c%28Office.15%29.aspx)|
|[Dialogs](http://msdn.microsoft.com/library/0d04aa87-9872-23e5-78e3-c9e3da2c8eb5%28Office.15%29.aspx)|
|[DisplayAlerts](http://msdn.microsoft.com/library/d9f36a99-e9c9-9a67-abaf-9c8e49b4febc%28Office.15%29.aspx)|
|[DisplayClipboardWindow](http://msdn.microsoft.com/library/16686caf-39ed-90fa-4a61-92b3f825cc6c%28Office.15%29.aspx)|
|[DisplayCommentIndicator](http://msdn.microsoft.com/library/8617da4e-97cb-fe57-bb51-a9c671e2ff27%28Office.15%29.aspx)|
|[DisplayDocumentActionTaskPane](http://msdn.microsoft.com/library/3b1fdce9-a6f1-ac6c-a14f-4ec8b35fd6a2%28Office.15%29.aspx)|
|[DisplayDocumentInformationPanel](http://msdn.microsoft.com/library/01810dbf-eab4-db5b-cb76-3196542f6e7b%28Office.15%29.aspx)|
|[DisplayExcel4Menus](http://msdn.microsoft.com/library/c281499a-cc84-5937-6436-78ecc8230a1d%28Office.15%29.aspx)|
|[DisplayFormulaAutoComplete](http://msdn.microsoft.com/library/bd6b78eb-2a5e-fbfa-e1f9-57810318f970%28Office.15%29.aspx)|
|[DisplayFormulaBar](http://msdn.microsoft.com/library/a54a313f-b416-5e5f-74d2-7435630b418e%28Office.15%29.aspx)|
|[DisplayFullScreen](http://msdn.microsoft.com/library/b42708ea-a273-c38a-5a61-d15e26c14fed%28Office.15%29.aspx)|
|[DisplayFunctionToolTips](http://msdn.microsoft.com/library/cc294f6d-3e81-9fdc-b758-0a581b03ba9c%28Office.15%29.aspx)|
|[DisplayInsertOptions](http://msdn.microsoft.com/library/81c1d837-463f-bc33-f815-7c6dc9678d1b%28Office.15%29.aspx)|
|[DisplayNoteIndicator](http://msdn.microsoft.com/library/96d43af3-0ceb-4bc2-ebaf-33cbe3e30a8a%28Office.15%29.aspx)|
|[DisplayPasteOptions](http://msdn.microsoft.com/library/da9cc6c1-e803-411a-220d-5c9c82d94504%28Office.15%29.aspx)|
|[DisplayRecentFiles](http://msdn.microsoft.com/library/006a38f4-11dd-1aad-8f5a-3771d4ab1ffc%28Office.15%29.aspx)|
|[DisplayScrollBars](http://msdn.microsoft.com/library/a81d2111-38eb-f156-28d7-a4abedf4839c%28Office.15%29.aspx)|
|[DisplayStatusBar](http://msdn.microsoft.com/library/bf70a679-bd50-cce7-0dc0-0dc57835038c%28Office.15%29.aspx)|
|[EditDirectlyInCell](http://msdn.microsoft.com/library/e867a786-5a34-2e12-e7c6-0637650611c0%28Office.15%29.aspx)|
|[EnableAnimations](http://msdn.microsoft.com/library/fb49fb3c-a842-73ab-1819-054f7403c85e%28Office.15%29.aspx)|
|[EnableAutoComplete](http://msdn.microsoft.com/library/eb5ccf8e-3e2d-2438-4dcf-d113cfdc3971%28Office.15%29.aspx)|
|[EnableCancelKey](http://msdn.microsoft.com/library/7c9c17b3-dd04-c914-4ed5-a6ef81ccf0c3%28Office.15%29.aspx)|
|[EnableCheckFileExtensions](http://msdn.microsoft.com/library/e518aec5-a261-47ba-a3fd-1da480c82612%28Office.15%29.aspx)|
|[EnableEvents](http://msdn.microsoft.com/library/5e14ce7b-02f6-03d4-2dfc-1df05a032301%28Office.15%29.aspx)|
|[EnableLargeOperationAlert](http://msdn.microsoft.com/library/c8454216-6e91-997a-566b-d00ca99e89a3%28Office.15%29.aspx)|
|[EnableLivePreview](http://msdn.microsoft.com/library/44163fba-3883-7744-de8b-36a0bd7f9e27%28Office.15%29.aspx)|
|[EnableMacroAnimations](http://msdn.microsoft.com/library/b1befccc-4f27-862b-8ab3-c862b5cb79b3%28Office.15%29.aspx)|
|[EnableSound](http://msdn.microsoft.com/library/8372b9dd-2929-6b5d-f51b-4409349dd6e6%28Office.15%29.aspx)|
|[ErrorCheckingOptions](http://msdn.microsoft.com/library/3821c6fd-e6c2-70cc-f546-70fdac6a6161%28Office.15%29.aspx)|
|[Excel4IntlMacroSheets](http://msdn.microsoft.com/library/39c70cd1-b54d-d781-d375-ca1d0715556f%28Office.15%29.aspx)|
|[Excel4MacroSheets](http://msdn.microsoft.com/library/d1ee907a-302c-4bd5-5455-75c328f94268%28Office.15%29.aspx)|
|[ExtendList](http://msdn.microsoft.com/library/b368047b-9d30-5a6f-a7db-748e3e91a3c0%28Office.15%29.aspx)|
|[FeatureInstall](http://msdn.microsoft.com/library/0bfe9d01-543c-9ea8-8ff6-2032f056b070%28Office.15%29.aspx)|
|[FileConverters](http://msdn.microsoft.com/library/7aebb0b3-6143-8dce-9893-e15decfe1c09%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/96a6fdc5-1bde-68dd-2493-9d8a92915afb%28Office.15%29.aspx)|
|[FileExportConverters](http://msdn.microsoft.com/library/1b7289ea-344f-cc3d-ec31-04d4196533ff%28Office.15%29.aspx)|
|[FileValidation](http://msdn.microsoft.com/library/6ec989d0-2ed8-b4d9-997c-4f91507e6fca%28Office.15%29.aspx)|
|[FileValidationPivot](http://msdn.microsoft.com/library/3cf6e177-9dbe-8ee8-3d84-599d7e2221da%28Office.15%29.aspx)|
|[FindFormat](http://msdn.microsoft.com/library/b2b62232-1f11-ec82-9344-edd39e0ae33d%28Office.15%29.aspx)|
|[FixedDecimal](http://msdn.microsoft.com/library/49b0a3de-bf5a-0130-e473-5b52f761932a%28Office.15%29.aspx)|
|[FixedDecimalPlaces](http://msdn.microsoft.com/library/e264dce3-4589-3e83-c931-5d69e3b8b3be%28Office.15%29.aspx)|
|[FlashFill](http://msdn.microsoft.com/library/85200392-3105-0bcd-a557-26e6a76fb5ac%28Office.15%29.aspx)|
|[FlashFillMode](http://msdn.microsoft.com/library/d77269c8-e47b-7d81-e5e4-68b0aa720a0d%28Office.15%29.aspx)|
|[FormulaBarHeight](http://msdn.microsoft.com/library/ff377046-06cb-9cf7-32f5-773da447c184%28Office.15%29.aspx)|
|[GenerateGetPivotData](http://msdn.microsoft.com/library/83effd5f-5101-ba1b-ab45-722e26074ea7%28Office.15%29.aspx)|
|[GenerateTableRefs](http://msdn.microsoft.com/library/3529fb4d-d311-6f92-9bf8-6b9f04d82ba8%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/2842f4c9-93b6-64a8-2394-72b47cf0cc83%28Office.15%29.aspx)|
|[HighQualityModeForGraphics](http://msdn.microsoft.com/library/7120b716-a0d4-e66e-2e98-4f2cd41324ef%28Office.15%29.aspx)|
|[Hinstance](http://msdn.microsoft.com/library/4551a0a2-0730-1288-7a13-b2beff2a2fca%28Office.15%29.aspx)|
|[HinstancePtr](http://msdn.microsoft.com/library/fddc40e9-08fc-34ef-60b2-41e8afa86575%28Office.15%29.aspx)|
|[Hwnd](http://msdn.microsoft.com/library/ed98b59c-1ebf-f319-f986-3406e4fdb766%28Office.15%29.aspx)|
|[IgnoreRemoteRequests](http://msdn.microsoft.com/library/94515401-eb26-a2d8-5013-33f1f38b884f%28Office.15%29.aspx)|
|[Interactive](http://msdn.microsoft.com/library/fe69429e-8715-770c-3e7a-c06a10a8e850%28Office.15%29.aspx)|
|[International](http://msdn.microsoft.com/library/e3849e31-a808-256c-4a94-c75c9d674d66%28Office.15%29.aspx)|
|[IsSandboxed](http://msdn.microsoft.com/library/d5a40aa3-470b-7861-691f-de418d13bd8b%28Office.15%29.aspx)|
|[Iteration](http://msdn.microsoft.com/library/51e5bd34-844b-3367-951a-6f2f8f9acf90%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/631879d9-f43f-4985-32d0-77bf314956eb%28Office.15%29.aspx)|
|[LargeOperationCellThousandCount](http://msdn.microsoft.com/library/f6140665-a5ec-bf17-c290-9e52686f605b%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/7a025afe-de39-26e7-5ac4-e6145ed2affd%28Office.15%29.aspx)|
|[LibraryPath](http://msdn.microsoft.com/library/783efa4a-640b-ab78-2831-da2ecd05558a%28Office.15%29.aspx)|
|[MailSession](http://msdn.microsoft.com/library/45dbbaa1-3da2-55f9-415b-ac9218d293dc%28Office.15%29.aspx)|
|[MailSystem](http://msdn.microsoft.com/library/df7b1238-bdf5-d9f8-9f50-585b489fd8a8%28Office.15%29.aspx)|
|[MapPaperSize](http://msdn.microsoft.com/library/c1d83fab-6abc-9103-78cf-047a503688b1%28Office.15%29.aspx)|
|[MathCoprocessorAvailable](http://msdn.microsoft.com/library/9424d6e1-f6f7-cc1b-7d20-987c8ed5e5a2%28Office.15%29.aspx)|
|[MaxChange](http://msdn.microsoft.com/library/5620bdff-d006-8c85-a1b8-1e3b31f21092%28Office.15%29.aspx)|
|[MaxIterations](http://msdn.microsoft.com/library/83f12597-9186-e415-a22b-9e028bd95169%28Office.15%29.aspx)|
|[MeasurementUnit](http://msdn.microsoft.com/library/2f48eda1-9d82-d8fc-ce89-2d33a4801e12%28Office.15%29.aspx)|
|[MergeInstances](http://msdn.microsoft.com/library/f406f2b2-802e-421c-9a80-f6f96a7b7c28%28Office.15%29.aspx)|
|[MouseAvailable](http://msdn.microsoft.com/library/b22f9d44-6a84-6716-d663-450f08c5557d%28Office.15%29.aspx)|
|[MoveAfterReturn](http://msdn.microsoft.com/library/9cdb96d5-e28a-b30c-25de-55a807d32c25%28Office.15%29.aspx)|
|[MoveAfterReturnDirection](http://msdn.microsoft.com/library/c11d8e36-755e-c911-de44-8b630b549418%28Office.15%29.aspx)|
|[MultiThreadedCalculation](http://msdn.microsoft.com/library/85aed55f-3127-6b4e-cc29-54bb0199d74d%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/f7fb2807-49de-c975-4931-ff825bfb0765%28Office.15%29.aspx)|
|[Names](http://msdn.microsoft.com/library/fe8727e4-3d04-47a1-13d2-386a7c68b5ed%28Office.15%29.aspx)|
|[NetworkTemplatesPath](http://msdn.microsoft.com/library/4710091a-a655-dd49-7ad8-0f4c64eda13a%28Office.15%29.aspx)|
|[NewWorkbook](http://msdn.microsoft.com/library/3a50a338-53be-3ac9-d398-c58084e19e6d%28Office.15%29.aspx)|
|[ODBCErrors](http://msdn.microsoft.com/library/47caef7a-fd3c-f67f-09c1-5ac21d65b67f%28Office.15%29.aspx)|
|[ODBCTimeout](http://msdn.microsoft.com/library/92262209-6a0f-f58f-e2d7-2f502f6bd397%28Office.15%29.aspx)|
|[OLEDBErrors](http://msdn.microsoft.com/library/0a42417f-f8b6-10bf-712a-44c1107f0f3e%28Office.15%29.aspx)|
|[OnWindow](http://msdn.microsoft.com/library/73ae5d34-66e6-3c1e-07f8-08850d13a4f5%28Office.15%29.aspx)|
|[OperatingSystem](http://msdn.microsoft.com/library/a36c5080-1d7e-a941-1bad-94f92522c7cf%28Office.15%29.aspx)|
|[OrganizationName](http://msdn.microsoft.com/library/4255a006-52df-66f6-2948-a9522e3adfef%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/e40a3599-1f4a-c79f-cc81-f629ecc888af%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/0ef5d0fc-f46a-c133-232a-8a20cf2d4034%28Office.15%29.aspx)|
|[PathSeparator](http://msdn.microsoft.com/library/573ef52d-3f1a-4fa3-4d4b-f047be67f26f%28Office.15%29.aspx)|
|[PivotTableSelection](http://msdn.microsoft.com/library/e0a93c11-2e2f-23af-6cad-b4f22883128e%28Office.15%29.aspx)|
|[PreviousSelections](http://msdn.microsoft.com/library/967ba122-700c-dca5-1b95-aeaf59e9f19c%28Office.15%29.aspx)|
|[PrintCommunication](http://msdn.microsoft.com/library/8b8ad1c5-1999-d733-44f4-734b7a388986%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/5fd20091-4c74-f39c-9842-a5a032048edd%28Office.15%29.aspx)|
|[PromptForSummaryInfo](http://msdn.microsoft.com/library/6a7799d7-327f-fdea-9c01-da48cf85655b%28Office.15%29.aspx)|
|[ProtectedViewWindows](http://msdn.microsoft.com/library/0f12ca56-f855-d05b-4a55-f31385a6489e%28Office.15%29.aspx)|
|[QuickAnalysis](http://msdn.microsoft.com/library/c79c04e7-0caf-470c-ee6d-dc613d6a4cf5%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/fec5050e-e6d9-6736-a9bc-b3e7d213a755%28Office.15%29.aspx)|
|[Ready](http://msdn.microsoft.com/library/4b9577ee-0f0c-dd0b-c1dd-90cde2c5fb1e%28Office.15%29.aspx)|
|[RecentFiles](http://msdn.microsoft.com/library/a64784af-4162-90fc-b955-963a1b1e747f%28Office.15%29.aspx)|
|[RecordRelative](http://msdn.microsoft.com/library/64e634e4-30e2-0794-1120-0960e32fe821%28Office.15%29.aspx)|
|[ReferenceStyle](http://msdn.microsoft.com/library/86c4931b-ab1a-0363-d048-5195707a952b%28Office.15%29.aspx)|
|[RegisteredFunctions](http://msdn.microsoft.com/library/c8922122-7de8-ebbb-0dfd-1dfe3974278e%28Office.15%29.aspx)|
|[ReplaceFormat](http://msdn.microsoft.com/library/df2242dc-9f23-b3c8-455d-1f0474eca873%28Office.15%29.aspx)|
|[RollZoom](http://msdn.microsoft.com/library/0bdad2a6-9d8d-cd69-cb73-45e9f92447d1%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/499f6045-1334-a8f8-9a04-f1aef7908312%28Office.15%29.aspx)|
|[RTD](http://msdn.microsoft.com/library/e181eb35-d8aa-4f46-3d50-6aa51776be7e%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/08fa0272-faeb-f8f2-c0f2-e001620cc838%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/f25b5608-035b-983a-545d-d720990c28be%28Office.15%29.aspx)|
|[Sheets](http://msdn.microsoft.com/library/729a512a-8faa-3a7e-758b-ff76e7200662%28Office.15%29.aspx)|
|[SheetsInNewWorkbook](http://msdn.microsoft.com/library/e2615d23-e0e0-34c4-0fd3-25f46a0d017b%28Office.15%29.aspx)|
|[ShowChartTipNames](http://msdn.microsoft.com/library/9f62fdc8-fcf0-eb4a-8ec4-d5d84cb96252%28Office.15%29.aspx)|
|[ShowChartTipValues](http://msdn.microsoft.com/library/886b2cf9-f6b3-3770-3082-28f2f99863cd%28Office.15%29.aspx)|
|[ShowDevTools](http://msdn.microsoft.com/library/de2c027f-cab2-f860-33aa-6c5fc63a5f73%28Office.15%29.aspx)|
|[ShowMenuFloaties](http://msdn.microsoft.com/library/8c0ac60a-e2cc-25f9-3915-f8c8ecd3690d%28Office.15%29.aspx)|
|[ShowQuickAnalysis](http://msdn.microsoft.com/library/043d9523-1fbc-0afd-2adf-9775e71058c0%28Office.15%29.aspx)|
|[ShowSelectionFloaties](http://msdn.microsoft.com/library/d2d74009-6b5e-ef62-2e32-83293b0f3f75%28Office.15%29.aspx)|
|[ShowStartupDialog](http://msdn.microsoft.com/library/8ea751c4-a4b1-a84a-9566-c4de8c5b9f67%28Office.15%29.aspx)|
|[ShowToolTips](http://msdn.microsoft.com/library/71293989-d0c4-f277-9d0b-c8fcda0ebf1f%28Office.15%29.aspx)|
|[SmartArtColors](http://msdn.microsoft.com/library/13f9961d-8bc2-b9b4-1c72-0cc74a4fc359%28Office.15%29.aspx)|
|[SmartArtLayouts](http://msdn.microsoft.com/library/00e0b95a-ba40-fb53-ebbe-4fd01b7a0e3a%28Office.15%29.aspx)|
|[SmartArtQuickStyles](http://msdn.microsoft.com/library/d4c9d4cf-b077-17b2-33dd-8449d0185b95%28Office.15%29.aspx)|
|[Speech](http://msdn.microsoft.com/library/981d5eef-55ff-54ee-a3ca-f009a6a575da%28Office.15%29.aspx)|
|[SpellingOptions](http://msdn.microsoft.com/library/c3d1970b-1276-9af7-88d6-e8e77bc32095%28Office.15%29.aspx)|
|[StandardFont](http://msdn.microsoft.com/library/6bde5ec0-8868-fa00-52e3-b7387f39f56d%28Office.15%29.aspx)|
|[StandardFontSize](http://msdn.microsoft.com/library/368ae001-7471-d104-573a-fc97d761f75e%28Office.15%29.aspx)|
|[StartupPath](http://msdn.microsoft.com/library/04bdd294-8127-37f2-7a39-b42923ac45b5%28Office.15%29.aspx)|
|[StatusBar](http://msdn.microsoft.com/library/91b043d7-b320-da4b-bdc7-3be1e1ffe3c6%28Office.15%29.aspx)|
|[TemplatesPath](http://msdn.microsoft.com/library/2db8397d-248b-7499-7959-1772b51d71a2%28Office.15%29.aspx)|
|[ThisCell](http://msdn.microsoft.com/library/83b9c009-7e01-4493-bda0-cd6246aba778%28Office.15%29.aspx)|
|[ThisWorkbook](http://msdn.microsoft.com/library/04b713dd-fd93-3cbc-f10b-05b9c3d107b1%28Office.15%29.aspx)|
|[ThousandsSeparator](http://msdn.microsoft.com/library/da244add-1c85-4636-2aff-b26feec215f3%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/309bf408-4f10-e343-228b-ebaad86d4b26%28Office.15%29.aspx)|
|[TransitionMenuKey](http://msdn.microsoft.com/library/3ea5b071-1ba7-19e9-1d6d-00bf128466e2%28Office.15%29.aspx)|
|[TransitionMenuKeyAction](http://msdn.microsoft.com/library/8f278d3b-9902-597a-9e4d-7f2fc3f22469%28Office.15%29.aspx)|
|[TransitionNavigKeys](http://msdn.microsoft.com/library/261afa51-44f7-4527-9145-b542cc68d812%28Office.15%29.aspx)|
|[UsableHeight](http://msdn.microsoft.com/library/536d2d03-0ce8-c28a-5a94-461fcfbd4ebf%28Office.15%29.aspx)|
|[UsableWidth](http://msdn.microsoft.com/library/b6c1cecb-28a5-8cdf-95ae-1b3b6e200dbb%28Office.15%29.aspx)|
|[UseClusterConnector](http://msdn.microsoft.com/library/9da42299-f23d-66e8-79b3-6105a0626db1%28Office.15%29.aspx)|
|[UsedObjects](http://msdn.microsoft.com/library/bf214478-990b-35c8-1e23-a9d1732e7ef3%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/fd55727d-8f79-14bf-038b-a31a56829a55%28Office.15%29.aspx)|
|[UserLibraryPath](http://msdn.microsoft.com/library/48e66da8-4db9-1262-9c0b-3a7f9f8e43ae%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/6cb2636c-ef3c-82fb-583d-8390cc815631%28Office.15%29.aspx)|
|[UseSystemSeparators](http://msdn.microsoft.com/library/eefa7bd0-9633-2f8a-cc80-61b1649fbace%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/298063f3-d2b3-ba55-7dcd-7419093094fb%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/e75dc57a-f9de-beb2-c50c-58445e47d63a%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/071cad0c-1cc0-8972-76f8-7c04d42765bd%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/4d702074-7d76-7b43-25e1-11d6a440392f%28Office.15%29.aspx)|
|[WarnOnFunctionNameConflict](http://msdn.microsoft.com/library/c29a9dbc-cd1f-18cc-2d44-ec639b0e67fa%28Office.15%29.aspx)|
|[Watches](http://msdn.microsoft.com/library/487c5cad-67bf-3bc9-dbc4-6bd8a105ed5e%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/eeb8ff27-d219-bade-3e0b-aed6e34d17d7%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/07e54620-c2f5-2354-f313-9756a0f17425%28Office.15%29.aspx)|
|[WindowsForPens](http://msdn.microsoft.com/library/798c0bd0-80f3-f6bd-a5d0-9abd88317bbc%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/f53d2bb8-b862-c55f-d9d5-68e705ca3415%28Office.15%29.aspx)|
|[Workbooks](http://msdn.microsoft.com/library/5291a324-87d7-3916-ffee-34c3389cea13%28Office.15%29.aspx)|
|[WorksheetFunction](http://msdn.microsoft.com/library/fd1333bf-8739-303d-30b4-85a99fb344b3%28Office.15%29.aspx)|
|[Worksheets](http://msdn.microsoft.com/library/ee9350d3-f24e-ed40-b267-8101d3267b4d%28Office.15%29.aspx)|


<br/>

## See also

- [Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
