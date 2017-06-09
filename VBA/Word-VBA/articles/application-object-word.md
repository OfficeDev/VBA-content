---
title: Application Object (Word)
keywords: vbawd10.chm2416
f1_keywords:
- vbawd10.chm2416
ms.prod: word
api_name:
- Word.Application
ms.assetid: d1cf6f8f-4e88-bf01-93b4-90a83f79cb44
ms.date: 06/08/2017
---


# Application Object (Word)

Represents the Microsoft Word application. The  **Application** object includes properties and methods that return top-level objects. For example, the **[ActiveDocument](http://msdn.microsoft.com/library/c20a7c9f-f8a4-7913-f53f-10baa6807def%28Office.15%29.aspx)** property returns a **[Document](document-object-word.md)** object.


## Remarks

Use the  **Application** property to return the **Application** object. The following example displays the user name for Word.


```
MsgBox Application.UserName
```

Many of the properties and methods that return the most common user-interface objects—such as the active document ( **ActiveDocument** property)—can be used without the **Application** object qualifier. For example, instead of writing `Application.ActiveDocument.PrintOut`, you can write  `ActiveDocument.PrintOut`. Properties and methods that can be used without the  **Application** object qualifier are considered "global." To view the global properties and methods in the **Object Browser**, click  **<globals>** at the top of the list in the **Classes** box. (Also see the **[Global](http://msdn.microsoft.com/library/b91e7459-08d5-ea8c-42e0-f7b9bfd1a72c%28Office.15%29.aspx)** object.)

Remarks

To use Automation (formerly OLE Automation) to control Word from another application, use the Microsoft Visual Basic  **CreateObject** or **GetObject** function to return a Word **Application** object. The following Microsoft Excel example starts Word (if it is not already running) and opens an existing document.




```
Set wrd = GetObject(, "Word.Application") 
wrd.Visible = True 
wrd.Documents.Open "C:\My Documents\Temp.doc" 
Set wrd = Nothing
```


## Events



|**Name**|
|:-----|
|[DocumentBeforeClose](http://msdn.microsoft.com/library/91c89b29-3110-85d7-c141-d1add3bb57f1%28Office.15%29.aspx)|
|[DocumentBeforePrint](http://msdn.microsoft.com/library/0736197a-5770-7e00-9882-86be0579c83e%28Office.15%29.aspx)|
|[DocumentBeforeSave](http://msdn.microsoft.com/library/cc1c6ec3-0e9e-5147-78a5-3a0c47fd5e90%28Office.15%29.aspx)|
|[DocumentChange](http://msdn.microsoft.com/library/853cbe7e-e4ac-2dba-826a-695d5c137d48%28Office.15%29.aspx)|
|[DocumentOpen](http://msdn.microsoft.com/library/21fdd3cd-8769-899e-5749-f64c0e15b265%28Office.15%29.aspx)|
|[DocumentSync](http://msdn.microsoft.com/library/9c83f692-8d05-2c52-11ef-46ac0ff69431%28Office.15%29.aspx)|
|[EPostageInsert](http://msdn.microsoft.com/library/33dfd513-7782-0877-7bf9-1b23cf995d4b%28Office.15%29.aspx)|
|[EPostageInsertEx](http://msdn.microsoft.com/library/494225b9-f55f-37d2-8ff0-086f8d917b05%28Office.15%29.aspx)|
|[EPostagePropertyDialog](http://msdn.microsoft.com/library/6d48fb9b-7826-2897-4deb-bde202fbd05b%28Office.15%29.aspx)|
|[MailMergeAfterMerge](http://msdn.microsoft.com/library/6eed8afa-efe6-0eba-6ab8-6c3ffc4e812d%28Office.15%29.aspx)|
|[MailMergeAfterRecordMerge](http://msdn.microsoft.com/library/6f483874-3999-815d-28b3-69fef89ed2be%28Office.15%29.aspx)|
|[MailMergeBeforeMerge](http://msdn.microsoft.com/library/968cf799-255f-b6fc-f576-7aec093ab1cb%28Office.15%29.aspx)|
|[MailMergeBeforeRecordMerge](http://msdn.microsoft.com/library/ce7b6c4f-b100-32eb-440c-c557f7dd7340%28Office.15%29.aspx)|
|[MailMergeDataSourceLoad](http://msdn.microsoft.com/library/56158dbd-45df-76ef-260d-117becd2e9ac%28Office.15%29.aspx)|
|[MailMergeDataSourceValidate](http://msdn.microsoft.com/library/31e03b87-b76c-9cfe-afb0-c9ee5cbcd13b%28Office.15%29.aspx)|
|[MailMergeDataSourceValidate2](http://msdn.microsoft.com/library/dba0dc60-a8c7-7e0c-ac02-4f5311534c89%28Office.15%29.aspx)|
|[MailMergeWizardSendToCustom](http://msdn.microsoft.com/library/b5dcd912-f1b5-96d6-3221-d294211b6611%28Office.15%29.aspx)|
|[MailMergeWizardStateChange](http://msdn.microsoft.com/library/d112d3f1-7fe7-1db6-891b-917598eea2ef%28Office.15%29.aspx)|
|[NewDocument](http://msdn.microsoft.com/library/afe5b924-3067-69e7-4331-a9ea2b30b9b5%28Office.15%29.aspx)|
|[ProtectedViewWindowActivate](http://msdn.microsoft.com/library/ae68e1aa-7cec-cd76-ee0e-71a051c5b6e3%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeClose](http://msdn.microsoft.com/library/4557dd3d-b795-94d9-ee53-5e83dcdd03d0%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeEdit](http://msdn.microsoft.com/library/1ea33944-1b2f-f914-f04a-81751cc750f8%28Office.15%29.aspx)|
|[ProtectedViewWindowDeactivate](http://msdn.microsoft.com/library/bd80056b-edce-7e0b-c61a-31ebda24a416%28Office.15%29.aspx)|
|[ProtectedViewWindowOpen](http://msdn.microsoft.com/library/42126a64-0227-d006-760e-ec11c59ef533%28Office.15%29.aspx)|
|[ProtectedViewWindowSize](http://msdn.microsoft.com/library/b28d53f9-783f-6d68-2080-a0b1d8484c43%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/3e05cf42-47c9-6a1b-b7da-09abe9746fd5%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/f1340e1e-6aec-edaa-78c2-47e3e1d5299f%28Office.15%29.aspx)|
|[WindowBeforeDoubleClick](http://msdn.microsoft.com/library/ece03591-0410-9dac-dedf-72c736dd477a%28Office.15%29.aspx)|
|[WindowBeforeRightClick](http://msdn.microsoft.com/library/c2d550e5-6781-a05f-41f6-eb9839aef208%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/70b86ecc-40ba-6e70-b430-4fce6083ff2d%28Office.15%29.aspx)|
|[WindowSelectionChange](http://msdn.microsoft.com/library/2c5cc640-a3a4-46b2-3352-20a057854b3a%28Office.15%29.aspx)|
|[WindowSize](http://msdn.microsoft.com/library/96d55786-52c8-68a9-b9e9-b29c320a435a%28Office.15%29.aspx)|
|[XMLSelectionChange](http://msdn.microsoft.com/library/a25d4e87-9b29-77b4-ddea-7692a0b56a8a%28Office.15%29.aspx)|
|[XMLValidationError](http://msdn.microsoft.com/library/bb75a555-fb5e-fb7b-f152-4c6436ecb1c7%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/7a6d6f71-cf69-1b96-10ed-2462ac6ff749%28Office.15%29.aspx)|
|[AddAddress](http://msdn.microsoft.com/library/9114f213-9e43-a65c-7513-631820481967%28Office.15%29.aspx)|
|[AutomaticChange](http://msdn.microsoft.com/library/40538590-c71c-aafb-4e3b-e8759cb0116c%28Office.15%29.aspx)|
|[BuildKeyCode](http://msdn.microsoft.com/library/0bbc515d-5f39-fed8-2b86-99979af928a9%28Office.15%29.aspx)|
|[CentimetersToPoints](http://msdn.microsoft.com/library/ca57a957-cc39-49ff-5e51-608e7985fd51%28Office.15%29.aspx)|
|[ChangeFileOpenDirectory](http://msdn.microsoft.com/library/9f044713-6e97-7219-8083-7d7d2cbb1b0f%28Office.15%29.aspx)|
|[CheckGrammar](http://msdn.microsoft.com/library/4675bda9-c31d-efdc-4def-38bfdeb200e4%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/88ea2134-cdbf-2bd5-bd6a-ff0c32a0f568%28Office.15%29.aspx)|
|[CleanString](http://msdn.microsoft.com/library/00fd8b33-77b0-d17a-b4f2-52b3892ed912%28Office.15%29.aspx)|
|[CompareDocuments](http://msdn.microsoft.com/library/511c811f-3f2b-9b93-f339-32324569a765%28Office.15%29.aspx)|
|[DDEExecute](http://msdn.microsoft.com/library/0f83607e-ba56-70d7-091e-411ec73fdfa7%28Office.15%29.aspx)|
|[DDEInitiate](http://msdn.microsoft.com/library/b9f7e0d9-e31f-370d-dec1-52721a4b712f%28Office.15%29.aspx)|
|[DDEPoke](http://msdn.microsoft.com/library/b782fc34-551f-288f-e087-5429f7ee7814%28Office.15%29.aspx)|
|[DDERequest](http://msdn.microsoft.com/library/beed4867-0e2d-15be-82ae-1aba11f0a21a%28Office.15%29.aspx)|
|[DDETerminate](http://msdn.microsoft.com/library/c469656c-edf8-3ce2-b09b-0883faba8943%28Office.15%29.aspx)|
|[DDETerminateAll](http://msdn.microsoft.com/library/1e8a0805-9bdd-add9-7184-533a0d2c5d9d%28Office.15%29.aspx)|
|[DefaultWebOptions](http://msdn.microsoft.com/library/ee683d3c-b331-cccd-27ec-b3258b42961e%28Office.15%29.aspx)|
|[GetAddress](http://msdn.microsoft.com/library/b0081a05-be87-d0e4-31a6-b0aab02a3371%28Office.15%29.aspx)|
|[GetDefaultTheme](http://msdn.microsoft.com/library/967760c0-4f99-5fae-026d-5ac60358d21c%28Office.15%29.aspx)|
|[GetSpellingSuggestions](http://msdn.microsoft.com/library/9ddf8aa8-10cc-8dd3-bc87-cdd5ccd214b5%28Office.15%29.aspx)|
|[GoBack](http://msdn.microsoft.com/library/d1113bc7-4ad3-f4da-0442-c11f5e22b2a8%28Office.15%29.aspx)|
|[GoForward](http://msdn.microsoft.com/library/5fb15a4f-d022-b82c-17ef-ec3808563a16%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/ff64e6bd-e29b-7cfc-437b-df8b8e59ce59%28Office.15%29.aspx)|
|[HelpTool](http://msdn.microsoft.com/library/81a93cf3-7695-8a35-d836-fc7a20622d92%28Office.15%29.aspx)|
|[InchesToPoints](http://msdn.microsoft.com/library/67a7e59c-bc61-be03-852d-05fadebef148%28Office.15%29.aspx)|
|[Keyboard](http://msdn.microsoft.com/library/67745d17-3dec-b4d9-919e-49925f2a7e34%28Office.15%29.aspx)|
|[KeyboardBidi](http://msdn.microsoft.com/library/5c54dcc0-8499-27f8-4b79-44764526d43b%28Office.15%29.aspx)|
|[KeyboardLatin](http://msdn.microsoft.com/library/115bb0af-470b-994f-f4f8-f714568ff469%28Office.15%29.aspx)|
|[KeyString](http://msdn.microsoft.com/library/20525053-3cf8-bdf8-cb67-cca39bf2b30c%28Office.15%29.aspx)|
|[LinesToPoints](http://msdn.microsoft.com/library/f146db0f-35f6-d25d-2674-e35a7c08801b%28Office.15%29.aspx)|
|[ListCommands](http://msdn.microsoft.com/library/425abd0f-c9c4-c4ab-b308-e7876ace5778%28Office.15%29.aspx)|
|[LoadMasterList](http://msdn.microsoft.com/library/f7722058-f097-3b8c-f124-df479e3efde6%28Office.15%29.aspx)|
|[LookupNameProperties](http://msdn.microsoft.com/library/e876b098-29c5-6302-97bb-c5f25ca4e101%28Office.15%29.aspx)|
|[MergeDocuments](http://msdn.microsoft.com/library/445fe4df-a41a-6be0-f646-d310c71ec25e%28Office.15%29.aspx)|
|[MillimetersToPoints](http://msdn.microsoft.com/library/13cf2786-709a-d473-0b6d-4fddabb465b8%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/030b6ae1-50bd-8d3e-e760-509c54a6e152%28Office.15%29.aspx)|
|[NewWindow](http://msdn.microsoft.com/library/0af15be1-7002-bd73-13da-19635d09b034%28Office.15%29.aspx)|
|[NextLetter](http://msdn.microsoft.com/library/b6b9e115-4efe-caf8-f0ee-53e41d8cd5d5%28Office.15%29.aspx)|
|[OnTime](http://msdn.microsoft.com/library/732d03cc-9dd6-5961-9763-048f72dea4d2%28Office.15%29.aspx)|
|[OrganizerCopy](http://msdn.microsoft.com/library/a23452aa-7372-ca58-291f-164e6000162d%28Office.15%29.aspx)|
|[OrganizerDelete](http://msdn.microsoft.com/library/45b394fc-cdd5-18ff-f30d-7339237a1b41%28Office.15%29.aspx)|
|[OrganizerRename](http://msdn.microsoft.com/library/abbe323c-b882-e497-608f-80004e166c8a%28Office.15%29.aspx)|
|[PicasToPoints](http://msdn.microsoft.com/library/ef812e9a-4bf5-b457-afa2-06371b411605%28Office.15%29.aspx)|
|[PixelsToPoints](http://msdn.microsoft.com/library/f5e2e3f2-1e58-d84f-c73a-f6414fa48c3d%28Office.15%29.aspx)|
|[PointsToCentimeters](http://msdn.microsoft.com/library/48ddcc3b-8b77-d363-4d94-aa64b74604a8%28Office.15%29.aspx)|
|[PointsToInches](http://msdn.microsoft.com/library/cabbab9b-ad04-576f-ee7c-921a38b4948c%28Office.15%29.aspx)|
|[PointsToLines](http://msdn.microsoft.com/library/8393f70f-4c2e-d74b-6add-f1d7f40ea75c%28Office.15%29.aspx)|
|[PointsToMillimeters](http://msdn.microsoft.com/library/b4933ff2-1d64-3ba2-8c28-69b7fa783358%28Office.15%29.aspx)|
|[PointsToPicas](http://msdn.microsoft.com/library/35d3f08b-bc4f-b65c-8b57-816146b37c77%28Office.15%29.aspx)|
|[PointsToPixels](http://msdn.microsoft.com/library/fc8eabb3-75f0-e456-bbd0-c17daa5ad1f3%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/f795218e-cd49-f3ac-c03d-9a9424361392%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/3913ee8b-291b-e81c-b106-01007738c7a0%28Office.15%29.aspx)|
|[PutFocusInMailHeader](http://msdn.microsoft.com/library/ca57a93b-1487-d19c-34c9-02484ce25485%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/0279d848-a8b7-dac7-1e84-a55c72789e3b%28Office.15%29.aspx)|
|[Repeat](http://msdn.microsoft.com/library/811e9f1c-cbdc-01dc-1e76-5521976943ed%28Office.15%29.aspx)|
|[ResetIgnoreAll](http://msdn.microsoft.com/library/8a6dcb30-23bb-70bb-e257-e519bc63a289%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/6614a0d8-eb2a-01fc-eeb6-4f8abc510bf8%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/d7d89a15-caea-d055-c60a-0e31acdfc2aa%28Office.15%29.aspx)|
|[ScreenRefresh](http://msdn.microsoft.com/library/303db23c-492c-5e33-0363-7ef6433dc90e%28Office.15%29.aspx)|
|[SetDefaultTheme](http://msdn.microsoft.com/library/7c51ff47-92d7-724f-0334-b789d2441313%28Office.15%29.aspx)|
|[ShowClipboard](http://msdn.microsoft.com/library/93696be3-3fd5-eb31-391c-d94e83d39d2b%28Office.15%29.aspx)|
|[ShowMe](http://msdn.microsoft.com/library/d9ebcfff-51dc-ac48-94ba-5cd99cc7373c%28Office.15%29.aspx)|
|[SubstituteFont](http://msdn.microsoft.com/library/2563bf9a-31ea-4104-b26b-538eb7e27f85%28Office.15%29.aspx)|
|[ToggleKeyboard](http://msdn.microsoft.com/library/a7af90f6-28e5-6655-ae5b-c01ed64da52f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveDocument](http://msdn.microsoft.com/library/c20a7c9f-f8a4-7913-f53f-10baa6807def%28Office.15%29.aspx)|
|[ActiveEncryptionSession](http://msdn.microsoft.com/library/a9ea5257-4cda-e25d-8af5-e29c10b7faa2%28Office.15%29.aspx)|
|[ActivePrinter](http://msdn.microsoft.com/library/835e350a-e069-e751-a7d7-1e9bb2483b4a%28Office.15%29.aspx)|
|[ActiveProtectedViewWindow](http://msdn.microsoft.com/library/2ba10f3d-3f43-5628-a5fc-3c65b290ef72%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/0a376a89-7bba-d5fd-8073-7cb2e79a87a8%28Office.15%29.aspx)|
|[AddIns](http://msdn.microsoft.com/library/8e464524-1304-6a4a-f2f0-5f652d5c8123%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/90d01c40-6b41-7b61-d989-6a864e6c2ca3%28Office.15%29.aspx)|
|[ArbitraryXMLSupportAvailable](http://msdn.microsoft.com/library/5cf53ae7-200b-811e-7946-4fefe825eaec%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/626b95a3-edae-977d-3c10-7a78fdcefeff%28Office.15%29.aspx)|
|[AutoCaptions](http://msdn.microsoft.com/library/6dd68657-3880-76eb-0dc4-91eb58fb0815%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/0a857e58-f37a-6023-fd13-bcb93109fdcd%28Office.15%29.aspx)|
|[AutoCorrectEmail](http://msdn.microsoft.com/library/20e94c20-ead7-f16f-b70f-c37d9f34a59e%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/2bc4f55c-d209-013b-77e4-ada7963bdee9%28Office.15%29.aspx)|
|[BackgroundPrintingStatus](http://msdn.microsoft.com/library/74fabdd0-55d8-63c6-4608-36af8138b3c1%28Office.15%29.aspx)|
|[BackgroundSavingStatus](http://msdn.microsoft.com/library/9cf29eb4-fc80-91ad-2867-6dc9d48e11c7%28Office.15%29.aspx)|
|[Bibliography](http://msdn.microsoft.com/library/25299c14-2198-9fde-0d5c-a6b96eda69ac%28Office.15%29.aspx)|
|[BrowseExtraFileTypes](http://msdn.microsoft.com/library/e411bb7a-d40f-1bda-5424-6202ba346717%28Office.15%29.aspx)|
|[Browser](http://msdn.microsoft.com/library/79b1967d-e661-8953-7bb2-a35eadbfae54%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/e22e7633-9327-eacc-3936-3d113381f675%28Office.15%29.aspx)|
|[CapsLock](http://msdn.microsoft.com/library/73cc2530-5178-d348-739e-c4605b8f207d%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/5554fa04-0744-400d-fd8c-2fe36d4ad9a3%28Office.15%29.aspx)|
|[CaptionLabels](http://msdn.microsoft.com/library/cf59346d-2ff5-938b-52ea-e2931422fd88%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/dea8365d-aadf-6667-ade0-2bef1622fd39%28Office.15%29.aspx)|
|[CheckLanguage](http://msdn.microsoft.com/library/25c2a119-2cae-48e4-1d54-cafc763b90fa%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/2e13bcfd-f2e1-61a5-164a-24cf121011a4%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/1082697d-edc8-c619-40d1-466d2ebf3817%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/6afdfc30-5021-7b09-a148-48db16d5efbd%28Office.15%29.aspx)|
|[CustomDictionaries](http://msdn.microsoft.com/library/1c6dca90-70f0-6b52-72d1-debda33d2ba0%28Office.15%29.aspx)|
|[CustomizationContext](http://msdn.microsoft.com/library/87c4fb87-1a59-fc0f-ca92-47e5d9c7c588%28Office.15%29.aspx)|
|[DefaultLegalBlackline](http://msdn.microsoft.com/library/a22afc29-1f7d-73af-75c2-7ce2fbe2250f%28Office.15%29.aspx)|
|[DefaultSaveFormat](http://msdn.microsoft.com/library/e15d8cc9-f6da-ccb0-784f-02fe9dc7ee6a%28Office.15%29.aspx)|
|[DefaultTableSeparator](http://msdn.microsoft.com/library/eb393e87-c408-8911-a1e3-8f04e5ce66c6%28Office.15%29.aspx)|
|[Dialogs](http://msdn.microsoft.com/library/17acdfab-32d2-ddb8-04aa-692f9ffb20b8%28Office.15%29.aspx)|
|[DisplayAlerts](http://msdn.microsoft.com/library/23d35e76-d5be-c1ed-4312-b6c220413882%28Office.15%29.aspx)|
|[DisplayAutoCompleteTips](http://msdn.microsoft.com/library/1ffcf473-d6f5-e2e7-c02c-0038b3fd3004%28Office.15%29.aspx)|
|[DisplayDocumentInformationPanel](http://msdn.microsoft.com/library/aa09e60f-6378-28b3-ff94-40f54a29d7b1%28Office.15%29.aspx)|
|[DisplayRecentFiles](http://msdn.microsoft.com/library/d8c96e18-7bbc-baa0-66ae-af91ee631a26%28Office.15%29.aspx)|
|[DisplayScreenTips](http://msdn.microsoft.com/library/07a03053-4973-27e2-6f0c-f67ff03c8bcf%28Office.15%29.aspx)|
|[DisplayScrollBars](http://msdn.microsoft.com/library/23b3957a-e4c1-b422-836a-074f84ff2f8e%28Office.15%29.aspx)|
|[Documents](http://msdn.microsoft.com/library/7e477cb3-ae65-685a-0083-1826efe86703%28Office.15%29.aspx)|
|[DontResetInsertionPointProperties](http://msdn.microsoft.com/library/3e6dfd03-9ab9-43c2-378c-0d97c69e14b3%28Office.15%29.aspx)|
|[EmailOptions](http://msdn.microsoft.com/library/28547346-6119-b763-339e-b04af1c8268f%28Office.15%29.aspx)|
|[EmailTemplate](http://msdn.microsoft.com/library/339a421e-b608-4063-a6e8-a08ba4debaf5%28Office.15%29.aspx)|
|[EnableCancelKey](http://msdn.microsoft.com/library/dd7d6885-7306-c6f3-56ff-e6f828adc4ea%28Office.15%29.aspx)|
|[FeatureInstall](http://msdn.microsoft.com/library/4abb8363-dee0-0209-2e56-54cea87fe692%28Office.15%29.aspx)|
|[FileConverters](http://msdn.microsoft.com/library/90f58ceb-6fb3-ee48-9637-effe39163888%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/ef478a81-db1d-4bf4-a146-3ff7dd84116b%28Office.15%29.aspx)|
|[FileValidation](http://msdn.microsoft.com/library/2f88d1a7-98a7-9ec6-09b3-a09c1a934e01%28Office.15%29.aspx)|
|[FindKey](http://msdn.microsoft.com/library/f648e9a5-626b-3923-46e4-a0c9c1dfc815%28Office.15%29.aspx)|
|[FocusInMailHeader](http://msdn.microsoft.com/library/fba9d08b-1950-b825-5f1a-14d671181b22%28Office.15%29.aspx)|
|[FontNames](http://msdn.microsoft.com/library/6aeadf51-79c7-1123-ea64-582ceee26443%28Office.15%29.aspx)|
|[HangulHanjaDictionaries](http://msdn.microsoft.com/library/453e2a77-f363-5afc-d9a3-26f8b6516b4c%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/070e8b38-8ed7-ff4a-0980-7ffc7cb901d4%28Office.15%29.aspx)|
|[International](http://msdn.microsoft.com/library/907c2908-01a6-2a83-9968-98c21b699f4b%28Office.15%29.aspx)|
|[IsObjectValid](http://msdn.microsoft.com/library/94cb08e4-2a4f-5ebf-25b8-6492e35f5695%28Office.15%29.aspx)|
|[IsSandboxed](http://msdn.microsoft.com/library/13fbfbda-b9e5-4f5d-46e2-2d8b157c77a1%28Office.15%29.aspx)|
|[KeyBindings](http://msdn.microsoft.com/library/68e08a9a-6547-f722-078e-b603b9f3e9cb%28Office.15%29.aspx)|
|[KeysBoundTo](http://msdn.microsoft.com/library/55967f9f-a2e0-eaae-a371-0fed82100138%28Office.15%29.aspx)|
|[LandscapeFontNames](http://msdn.microsoft.com/library/59599ca0-0c6f-8d4a-9f4e-e98c5c241944%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/b24f0861-1f7a-ecd9-7084-39c65df4fcc3%28Office.15%29.aspx)|
|[Languages](http://msdn.microsoft.com/library/f81cfcb6-33e2-bb8e-2ac4-b4f9df833946%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/62bd3b7e-e9b4-3158-4531-4dfffd9cdb02%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/de6b783d-5001-0fed-33e8-15f9d338986c%28Office.15%29.aspx)|
|[ListGalleries](http://msdn.microsoft.com/library/769d3494-3fc3-5a4b-e6d1-a3910107c8bd%28Office.15%29.aspx)|
|[MacroContainer](http://msdn.microsoft.com/library/9c2d37b8-d5c3-d13b-3bf9-54e1352b1855%28Office.15%29.aspx)|
|[MailingLabel](http://msdn.microsoft.com/library/7eba3273-4a4c-6cdf-004a-4a0d214d6127%28Office.15%29.aspx)|
|[MailMessage](http://msdn.microsoft.com/library/82bca039-0b6b-4489-27bf-18746dc639d2%28Office.15%29.aspx)|
|[MailSystem](http://msdn.microsoft.com/library/d8f97baa-2c50-c2af-0e50-f2de5d017b62%28Office.15%29.aspx)|
|[MAPIAvailable](http://msdn.microsoft.com/library/2cb2fc8c-1ef6-98b8-fa72-0705637ad3ac%28Office.15%29.aspx)|
|[MathCoprocessorAvailable](http://msdn.microsoft.com/library/207b7f3f-4113-7069-51e3-10658ec3654f%28Office.15%29.aspx)|
|[MouseAvailable](http://msdn.microsoft.com/library/25ad78ad-c267-35ec-9124-0496c034fa50%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/079525c4-1817-e142-de08-86a72bd18c55%28Office.15%29.aspx)|
|[NewDocument](http://msdn.microsoft.com/library/2f68f98e-1aad-eeac-59c7-4cd5f9d7ad6a%28Office.15%29.aspx)|
|[NormalTemplate](http://msdn.microsoft.com/library/0ffd1cfd-78da-5189-2504-bbc04bf5b484%28Office.15%29.aspx)|
|[NumLock](http://msdn.microsoft.com/library/0c20c000-2df9-1483-91be-cacf1abe0ff0%28Office.15%29.aspx)|
|[OMathAutoCorrect](http://msdn.microsoft.com/library/babed2d9-eecf-de72-a1f2-9387d068e74a%28Office.15%29.aspx)|
|[OpenAttachmentsInFullScreen](http://msdn.microsoft.com/library/295af420-fbe0-7753-2f7f-afabb5f0818c%28Office.15%29.aspx)|
|[Options](http://msdn.microsoft.com/library/87bf2092-8707-d375-d4d6-f7420be1fe7d%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1855987d-a710-5919-9fec-a53c24a2ef5e%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/224b4c66-f49c-55f1-8b6b-74f5ed979a3d%28Office.15%29.aspx)|
|[PathSeparator](http://msdn.microsoft.com/library/29347a13-8edb-0b02-32c3-d091eb52c9f1%28Office.15%29.aspx)|
|[PickerDialog](http://msdn.microsoft.com/library/1d3de67e-6760-1b5d-9ff6-fccb2de32455%28Office.15%29.aspx)|
|[PortraitFontNames](http://msdn.microsoft.com/library/21c3802b-43ad-3d8f-34de-af9af4d29bcf%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/6f522dc1-60ad-d5c1-029b-961fce1992e5%28Office.15%29.aspx)|
|[ProtectedViewWindows](http://msdn.microsoft.com/library/eb1c8cae-c0da-0a84-316e-808302869b26%28Office.15%29.aspx)|
|[RecentFiles](http://msdn.microsoft.com/library/517fb0cf-2dfb-f0a0-0882-f233198768d6%28Office.15%29.aspx)|
|[RestrictLinkedStyles](http://msdn.microsoft.com/library/0d2033bc-9cf4-1f57-a9c7-56eaf0a55257%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/660284d1-2b00-eba0-28bc-36bc6400fcf4%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/d2362378-06a1-3a1a-2bd0-358f190eb6f3%28Office.15%29.aspx)|
|[ShowAnimation](http://msdn.microsoft.com/library/e58b9f65-2fe5-4c88-39ab-ae7aa77490b1%28Office.15%29.aspx)|
|[ShowStartupDialog](http://msdn.microsoft.com/library/ecee5bb2-271b-f1fc-c25c-a77a59f5df03%28Office.15%29.aspx)|
|[ShowStylePreviews](http://msdn.microsoft.com/library/16edc0cd-29a4-f951-8344-c4603fc047f7%28Office.15%29.aspx)|
|[ShowVisualBasicEditor](http://msdn.microsoft.com/library/eb0a9d3f-3eba-f7fb-2939-a7274744b4b8%28Office.15%29.aspx)|
|[SmartArtColors](http://msdn.microsoft.com/library/e2cb12c4-3162-2327-9210-bd912dffa8e9%28Office.15%29.aspx)|
|[SmartArtLayouts](http://msdn.microsoft.com/library/dcbaf620-0865-8f2f-ef97-456edd0d70e3%28Office.15%29.aspx)|
|[SmartArtQuickStyles](http://msdn.microsoft.com/library/47cca923-fc88-6973-926c-2fa69c2f0f10%28Office.15%29.aspx)|
|[SpecialMode](http://msdn.microsoft.com/library/aa60d4dc-4abe-e461-12c9-fc8e890536ca%28Office.15%29.aspx)|
|[StartupPath](http://msdn.microsoft.com/library/1b73f234-358b-a360-ca69-ed00e0817038%28Office.15%29.aspx)|
|[StatusBar](http://msdn.microsoft.com/library/09ee8f1a-65e1-a852-9df1-660436a7bf72%28Office.15%29.aspx)|
|[SynonymInfo](http://msdn.microsoft.com/library/7aff62c5-d962-23b5-0e86-ae566f72914c%28Office.15%29.aspx)|
|[System](http://msdn.microsoft.com/library/871f3821-4e17-1c63-9b4b-1d4e2bfc97d5%28Office.15%29.aspx)|
|[TaskPanes](http://msdn.microsoft.com/library/0b0add9d-6c76-9dca-e7a5-3f653f5d1581%28Office.15%29.aspx)|
|[Tasks](http://msdn.microsoft.com/library/e78150fd-8c79-948b-7a46-5e560194aa48%28Office.15%29.aspx)|
|[Templates](http://msdn.microsoft.com/library/816e50d1-32b9-c8ff-6d2c-ad1113c952fc%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/bbce9fe2-8390-f73d-8fca-bd047df468be%28Office.15%29.aspx)|
|[UndoRecord](http://msdn.microsoft.com/library/d21c7089-2cdc-3d04-1073-ada649f21576%28Office.15%29.aspx)|
|[UsableHeight](http://msdn.microsoft.com/library/9723b59d-c5fe-8f39-8f0c-bdd209b7ae9a%28Office.15%29.aspx)|
|[UsableWidth](http://msdn.microsoft.com/library/e5ea64f7-2a7a-fdaa-20ff-fdf6196de874%28Office.15%29.aspx)|
|[UserAddress](http://msdn.microsoft.com/library/34f9bf48-8af4-4017-a648-13ab7612ca4a%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/65cf8ccc-f846-cd86-a8d6-0b2951bad604%28Office.15%29.aspx)|
|[UserInitials](http://msdn.microsoft.com/library/00f7d562-4ce5-00e1-bf9d-4325d47947b2%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/96f5ffb6-a20d-96f0-e3a4-0ad2dd47bf99%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/641109fd-7ece-9efd-65ba-56e223d8249c%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/7bdd0acc-1ed0-677c-f973-99a9199e030b%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/e3de4b88-55e4-cb6d-130f-9aea11f3eff6%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/ac9b369e-6661-ef67-6674-85ab02ef4621%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/860d9e12-4c02-be1f-64a7-ef0305881854%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/ae457f42-9c12-d0f4-e74e-d01610b9b4af%28Office.15%29.aspx)|
|[WordBasic](http://msdn.microsoft.com/library/8c405ea6-0073-f994-42b2-cacb986f1f1f%28Office.15%29.aspx)|
|[XMLNamespaces](http://msdn.microsoft.com/library/e7eac332-f805-5ceb-682c-482565ff0786%28Office.15%29.aspx)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

