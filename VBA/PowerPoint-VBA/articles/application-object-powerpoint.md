---
title: Application Object (PowerPoint)
keywords: vbapp10.chm504000
f1_keywords:
- vbapp10.chm504000
ms.prod: powerpoint
api_name:
- PowerPoint.Application
ms.assetid: 978c2b99-4271-b953-4283-73b5f3d96f41
ms.date: 06/08/2017
---


# Application Object (PowerPoint)

Represents the entire Microsoft PowerPoint application. 


## Remarks

The  **Application** object contains:


- Application-wide settings and options (the name of the active printer, for example).
    
- Properties that return top-level objects, such as  **ActivePresentation**, and **Windows**.
    


When you are writing code that will run from PowerPoint, you can use the following properties of the  **Application** object without the object qualifier: **ActivePresentation**, **ActiveWindow**, **AddIns**, **Presentations**, **SlideShowWindows**, **Windows**.

For example, instead of writing  `Application.ActiveWindow.Height = 200`, you can write  `ActiveWindow.Height = 200`.


## Example

Use the  **Application** property to return the **Application** object. The following example returns the path to the program file.


```
Dim MyPath As String

MyPath = Application.Path
```

The following example creates a PowerPoint  **Application** object in another application, starts PowerPoint (if it is not already running), and opens an existing presentation named "Ex_a2a.ppt".




```
Set ppt = New Powerpoint.Application

ppt.Visible = True

ppt.Presentations.Open "c:\My Documents\ex_a2a.ppt"
```


## Events



|**Name**|
|:-----|
|[AfterDragDropOnSlide](http://msdn.microsoft.com/library/1de9f2a4-565b-152a-452a-cb0c1a135c35%28Office.15%29.aspx)|
|[AfterNewPresentation](http://msdn.microsoft.com/library/d95bb247-2ebd-263f-d6b5-9918204b9130%28Office.15%29.aspx)|
|[AfterPresentationOpen](http://msdn.microsoft.com/library/3f783486-0ceb-166d-017b-0a41bd15cfa6%28Office.15%29.aspx)|
|[AfterShapeSizeChange](http://msdn.microsoft.com/library/0c7eacc9-445a-b1ec-1f48-6d11fbb842e9%28Office.15%29.aspx)|
|[ColorSchemeChanged](http://msdn.microsoft.com/library/8b517ce7-879d-bb96-477b-072477c991d5%28Office.15%29.aspx)|
|[NewPresentation](http://msdn.microsoft.com/library/63a6a83d-74c4-88ac-4972-d54907f5af8a%28Office.15%29.aspx)|
|[PresentationBeforeClose](http://msdn.microsoft.com/library/8c2d820b-aa44-287b-10ad-1dc6f4122231%28Office.15%29.aspx)|
|[PresentationBeforeSave](http://msdn.microsoft.com/library/40943fe2-796f-45db-db0d-44b66854e196%28Office.15%29.aspx)|
|[PresentationClose](http://msdn.microsoft.com/library/4057b50a-5f2d-78bf-d55a-d0781da27ea7%28Office.15%29.aspx)|
|[PresentationCloseFinal](http://msdn.microsoft.com/library/4972c700-9d7a-e43e-1e22-f9882368741e%28Office.15%29.aspx)|
|[PresentationNewSlide](http://msdn.microsoft.com/library/e9718cad-6411-d013-6c93-0370aa71a8f2%28Office.15%29.aspx)|
|[PresentationOpen](http://msdn.microsoft.com/library/1739cee9-cfc1-0650-de24-be699bafe910%28Office.15%29.aspx)|
|[PresentationPrint](http://msdn.microsoft.com/library/41a420b7-c5db-7869-6763-da9cec710d83%28Office.15%29.aspx)|
|[PresentationSave](http://msdn.microsoft.com/library/229a02a7-58e4-2445-3bd5-963e88438d7e%28Office.15%29.aspx)|
|[PresentationSync](http://msdn.microsoft.com/library/391b486e-7e92-bc90-224a-77c499cdf774%28Office.15%29.aspx)|
|[ProtectedViewWindowActivate](http://msdn.microsoft.com/library/3a7b3842-9524-9e42-b2b1-aff45e17d965%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeClose](http://msdn.microsoft.com/library/e10ffe16-aad8-1e2d-fd75-82243a56ef05%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeEdit](http://msdn.microsoft.com/library/8cfd38bf-8336-0106-a170-1319bcea0eb8%28Office.15%29.aspx)|
|[ProtectedViewWindowDeactivate](http://msdn.microsoft.com/library/c8d647f3-2f45-7811-9f99-d37c3c999c60%28Office.15%29.aspx)|
|[ProtectedViewWindowOpen](http://msdn.microsoft.com/library/642a0f98-7ff9-daea-33ad-a893a65b9782%28Office.15%29.aspx)|
|[SlideSelectionChanged](http://msdn.microsoft.com/library/a7bbdc4c-31e3-2072-8590-bced8bff6517%28Office.15%29.aspx)|
|[SlideShowBegin](http://msdn.microsoft.com/library/f70ca9cb-11a7-2a81-19bb-36e0b0ca0b97%28Office.15%29.aspx)|
|[SlideShowEnd](http://msdn.microsoft.com/library/e46f8177-e00b-6704-1606-dbf9e96bf812%28Office.15%29.aspx)|
|[SlideShowNextBuild](http://msdn.microsoft.com/library/63919ea5-57e4-853a-0e5a-94e1126cbfbf%28Office.15%29.aspx)|
|[SlideShowNextClick](http://msdn.microsoft.com/library/95a83383-62a4-a99b-3cd4-a69700bfbc3a%28Office.15%29.aspx)|
|[SlideShowNextSlide](http://msdn.microsoft.com/library/a73d051e-9f53-43bd-1f41-b9111197e464%28Office.15%29.aspx)|
|[SlideShowOnNext](http://msdn.microsoft.com/library/de72c6d6-0794-ad1d-5b25-478caaafd099%28Office.15%29.aspx)|
|[SlideShowOnPrevious](http://msdn.microsoft.com/library/466a5363-047b-f107-011b-6450db6a5f31%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/0d83fda3-b0ad-18df-57bf-c34dafcf782f%28Office.15%29.aspx)|
|[WindowBeforeDoubleClick](http://msdn.microsoft.com/library/9b270238-1658-df56-4208-9cb98666519c%28Office.15%29.aspx)|
|[WindowBeforeRightClick](http://msdn.microsoft.com/library/e6239915-f487-3619-c84f-d436d645e6c0%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/89bf2c09-a1a8-ed7f-74d5-49f8f7c027a7%28Office.15%29.aspx)|
|[WindowSelectionChange](http://msdn.microsoft.com/library/069f4afe-2302-28fa-4d86-57afe8c3c2ab%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/63a64e28-8e27-12b3-0189-4b6e5513bc00%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/97dabc76-1987-6e08-ea42-6762be6b7d60%28Office.15%29.aspx)|
|[OpenThemeFile](http://msdn.microsoft.com/library/b34d5a6f-8cf8-ce6a-3c0c-c1ed43c413c6%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/d7040179-ca03-563f-5bd9-80a5fd5e5d4b%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/21b8a0c4-10c8-d8c3-9214-adffad35f7d4%28Office.15%29.aspx)|
|[StartNewUndoEntry](http://msdn.microsoft.com/library/7f4f2236-6e6a-11e9-20b5-0fca5c126330%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Active](http://msdn.microsoft.com/library/94eb9039-ac4a-b8e0-dc66-c508521e3604%28Office.15%29.aspx)|
|[ActiveEncryptionSession](http://msdn.microsoft.com/library/73a174d5-a088-97d0-5f71-931456493224%28Office.15%29.aspx)|
|[ActivePresentation](http://msdn.microsoft.com/library/55ff4906-09e5-2c5c-0ed7-5f7a767542f7%28Office.15%29.aspx)|
|[ActivePrinter](http://msdn.microsoft.com/library/48ba3853-6a8f-d523-807a-8324e59adbb7%28Office.15%29.aspx)|
|[ActiveProtectedViewWindow](http://msdn.microsoft.com/library/c0a7e748-d7fc-4a63-62b8-0eed5cf1c5b5%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/762c1c6a-1f8a-f47a-7b75-006c745caee0%28Office.15%29.aspx)|
|[AddIns](http://msdn.microsoft.com/library/5a5a030f-45cd-3b82-f41a-eab53b1ed48f%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/0062855c-0756-b8fd-943e-e8f9297c9759%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/490fc728-c639-2a32-22b8-1757c14e9bd7%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/942341fe-5290-2903-db70-4e7cff0d75c7%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/e485e2f1-835c-33aa-c585-32fbd3af4a88%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/f6281931-8a78-9e8b-0a41-ae7d63f8755e%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/c31b3771-d7b1-7559-4480-75f91f1d1f52%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/f24029c9-f839-e9a4-d661-5f1e22080d46%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/3ba8a827-f585-b4f5-4ba0-20a0d791216c%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/3caec137-72b5-6ec9-3b79-acd55df62a3e%28Office.15%29.aspx)|
|[DisplayAlerts](http://msdn.microsoft.com/library/e18cf1f5-c456-8cd5-40e7-eec69c40811d%28Office.15%29.aspx)|
|[DisplayDocumentInformationPanel](http://msdn.microsoft.com/library/473f5e46-2615-b456-12ca-440afda0e642%28Office.15%29.aspx)|
|[DisplayGridLines](http://msdn.microsoft.com/library/b639cd4f-26d4-4f63-2fe0-18807bdeefa5%28Office.15%29.aspx)|
|[DisplayGuides](http://msdn.microsoft.com/library/637488b3-c657-6a78-d897-cb58122d80b2%28Office.15%29.aspx)|
|[FeatureInstall](http://msdn.microsoft.com/library/254fc432-9ee5-d978-19ac-5fa6f94daa94%28Office.15%29.aspx)|
|[FileConverters](http://msdn.microsoft.com/library/2eaa06eb-e32c-cf07-03a2-880048468188%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/0f0d5b6c-e478-6d15-7218-be04df978d6b%28Office.15%29.aspx)|
|[FileValidation](http://msdn.microsoft.com/library/90cc8bff-df3b-7a57-adcc-bbfb9c677468%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/4236df34-3381-2a36-9b51-05a28308377e%28Office.15%29.aspx)|
|[IsSandboxed](http://msdn.microsoft.com/library/c17eed5c-8612-5cd8-3ef6-a745d54d2a10%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/9603b5ed-2143-10f7-399b-2757b71c0525%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/8513a292-b293-19ec-18ce-0b444b8b4715%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/c7a59327-774a-8c55-17b4-053ae76bd623%28Office.15%29.aspx)|
|[NewPresentation](http://msdn.microsoft.com/library/9685db30-9d73-19ad-432b-8d79b2d6ee50%28Office.15%29.aspx)|
|[OperatingSystem](http://msdn.microsoft.com/library/5532197a-f6c3-825a-6492-e1c85d97a9d2%28Office.15%29.aspx)|
|[Options](http://msdn.microsoft.com/library/4f890917-68bc-bb02-914d-52ea8a82bbcb%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/aae10b96-e0e4-d055-f398-d26f4cab572d%28Office.15%29.aspx)|
|[Presentations](http://msdn.microsoft.com/library/d6f5f565-d593-e230-c3b9-2302bdd83644%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/27376e9f-47c6-7373-af34-4ce71723e6a6%28Office.15%29.aspx)|
|[ProtectedViewWindows](http://msdn.microsoft.com/library/21ffdddc-9e29-94ee-425d-c83d49dcf457%28Office.15%29.aspx)|
|[ShowStartupDialog](http://msdn.microsoft.com/library/acbd2597-c835-e285-e52c-5c86349d3199%28Office.15%29.aspx)|
|[ShowWindowsInTaskbar](http://msdn.microsoft.com/library/ad386fe5-9985-a1cc-cc52-1552bc12cad4%28Office.15%29.aspx)|
|[SlideShowWindows](http://msdn.microsoft.com/library/4beed51c-bb67-6208-c2b1-f1d5b6425d9b%28Office.15%29.aspx)|
|[SmartArtColors](http://msdn.microsoft.com/library/6a324540-8703-6e18-938d-b275e1f71610%28Office.15%29.aspx)|
|[SmartArtLayouts](http://msdn.microsoft.com/library/0b175f36-6333-f073-2545-abd342492ea1%28Office.15%29.aspx)|
|[SmartArtQuickStyles](http://msdn.microsoft.com/library/79fc3e91-0862-c294-dc0b-fe06d9c2c006%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/d8c70fc9-e0f1-ed53-7a22-150838599719%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/33a3d113-31f6-3705-cdb9-d5e07fa82820%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/c76b1e7e-db29-0ef8-fefb-9333b8350de0%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/8c28f542-56b2-49e3-8b77-a7424e00c773%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/ba9c122d-4283-1865-63f1-07bf746f1606%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/c6d001c6-b589-47bc-bf6a-d1cf9b277f3d%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/128f7da4-3cc3-1cda-6298-8bbc0b39a25c%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
