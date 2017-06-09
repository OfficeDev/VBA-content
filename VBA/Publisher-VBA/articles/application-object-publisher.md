---
title: Application Object (Publisher)
keywords: vbapb10.chm536936447
f1_keywords:
- vbapb10.chm536936447
ms.prod: publisher
api_name:
- Publisher.Application
ms.assetid: acfc7efb-e6a5-a89a-3aee-3cb4af2f3508
ms.date: 06/08/2017
---


# Application Object (Publisher)

Represents the Microsoft Publisher application. The  **Application** object includes properties and methods that return top-level objects. For example, the **ActiveDocument** property returns a **Document** object.


## Remarks

When using Microsoft Visual Basic for Applications in Publisher, all of the properties and methods of the  **Application** object can be used without the **Application** object qualifier. For example, instead of typing `Application.ActiveDocument.PrintOut`, you can type  `ActiveDocument.PrintOut`. Properties and methods that can be used without the  **Application** object qualifier are considered "global." To view the global properties and methods in the Object Browser, click **<globals>** at the top of the list in the **Classes** box. When accessing the Publisher object model from a non-Publisher project, all properties and methods must be fully qualified.


## Example

Use the  **[Application](http://msdn.microsoft.com/library/f3ed5997-b8ef-4729-4537-ae21424d2007%28Office.15%29.aspx)** property to return the **Application** object. The following example displays the application name.


```
Sub ShowAppName() 
 MsgBox Application.Name 
End Sub
```


## Events



|**Name**|
|:-----|
|[AfterPrint](http://msdn.microsoft.com/library/ddd5a1a4-8130-9e75-039c-e069a37390e8%28Office.15%29.aspx)|
|[BeforePrint](http://msdn.microsoft.com/library/4d819aab-726e-ab00-89e0-aedcb62d834e%28Office.15%29.aspx)|
|[DocumentBeforeClose](http://msdn.microsoft.com/library/d3ca4397-4df3-dc77-b758-d47e0bf13fe5%28Office.15%29.aspx)|
|[DocumentOpen](http://msdn.microsoft.com/library/3bdd4b38-ec40-a08f-3742-f81a6ed333b3%28Office.15%29.aspx)|
|[HideCatalogUI](http://msdn.microsoft.com/library/a7ac7594-18fe-355e-d270-d205c405862a%28Office.15%29.aspx)|
|[MailMergeAfterMerge](http://msdn.microsoft.com/library/dd01d8f5-f95e-e833-bb8b-708ced54240c%28Office.15%29.aspx)|
|[MailMergeAfterRecordMerge](http://msdn.microsoft.com/library/550c3310-01ba-718f-4c1d-cbf3ce077d27%28Office.15%29.aspx)|
|[MailMergeBeforeMerge](http://msdn.microsoft.com/library/735ef282-e99f-b3f2-c509-b180bea30d36%28Office.15%29.aspx)|
|[MailMergeBeforeRecordMerge](http://msdn.microsoft.com/library/67ae8255-336d-0ff8-7927-fbd31262c115%28Office.15%29.aspx)|
|[MailMergeDataSourceLoad](http://msdn.microsoft.com/library/afca3a05-d6a6-15f1-8cbf-593777066757%28Office.15%29.aspx)|
|[MailMergeDataSourceValidate](http://msdn.microsoft.com/library/8e18b0a0-8fe8-f72e-8a75-1585367cc796%28Office.15%29.aspx)|
|[MailMergeGenerateBarcode](http://msdn.microsoft.com/library/5da4ec65-32b6-ea05-09ad-d2224eafee30%28Office.15%29.aspx)|
|[MailMergeInsertBarcode](http://msdn.microsoft.com/library/6b901953-eaff-0189-1d33-678e935a2f7e%28Office.15%29.aspx)|
|[MailMergeRecipientListClose](http://msdn.microsoft.com/library/4fb77771-9897-8623-f4e7-61f631f04922%28Office.15%29.aspx)|
|[MailMergeWizardFollowUpCustom](http://msdn.microsoft.com/library/ac8cb695-69a4-83f7-8e13-66762f52f611%28Office.15%29.aspx)|
|[MailMergeWizardStateChange](http://msdn.microsoft.com/library/3d3fcdaa-af51-0a28-ff25-f2b92deceaf6%28Office.15%29.aspx)|
|[NewDocument](http://msdn.microsoft.com/library/629cf55c-5134-4207-14df-143b517b9f36%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/79948040-4848-b8e7-a70c-d23c1f416bac%28Office.15%29.aspx)|
|[ShowCatalogUI](http://msdn.microsoft.com/library/8a5a3798-4b95-d77f-70f6-d69dd9dc8f99%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/a7e4e396-9661-763c-8e41-dc279757af94%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/84473784-7c03-4c9e-3e1b-9bf6ec7e1fbc%28Office.15%29.aspx)|
|[WindowPageChange](http://msdn.microsoft.com/library/bb636f6e-da4b-7271-9f59-2b7000270c16%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[CentimetersToPoints](http://msdn.microsoft.com/library/6eda6692-ea9a-c4ad-6991-066fdc23bd2c%28Office.15%29.aspx)|
|[ChangeFileOpenDirectory](http://msdn.microsoft.com/library/9178881c-2f7f-9063-31d1-14d4745f0666%28Office.15%29.aspx)|
|[EmusToPoints](http://msdn.microsoft.com/library/941e5975-ca7a-38dc-8116-e90b2a2ab6e5%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/37b51399-5897-4003-a0a9-9829a8adf8ed%28Office.15%29.aspx)|
|[InchesToPoints](http://msdn.microsoft.com/library/32c8740f-ad14-c947-b960-500378a5873d%28Office.15%29.aspx)|
|[IsValidObject](http://msdn.microsoft.com/library/56b2bc3a-3e8e-058c-046a-146f0fbb294a%28Office.15%29.aspx)|
|[LinesToPoints](http://msdn.microsoft.com/library/55c531aa-5619-6f7f-54e7-7721cb70640e%28Office.15%29.aspx)|
|[MillimetersToPoints](http://msdn.microsoft.com/library/40ec9abd-cc1e-9f44-3312-d6689b4822e4%28Office.15%29.aspx)|
|[NewDocument](http://msdn.microsoft.com/library/9beb6176-0c46-0ba0-8d41-a9021c624223%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/560ac406-f058-8fd8-4b6d-978ff369de9b%28Office.15%29.aspx)|
|[PicasToPoints](http://msdn.microsoft.com/library/64d3e435-dcc1-d637-7aac-cc9a9bf81e76%28Office.15%29.aspx)|
|[PixelsToPoints](http://msdn.microsoft.com/library/5d7e453f-e962-e557-48e4-44766d0c64d9%28Office.15%29.aspx)|
|[PointsToCentimeters](http://msdn.microsoft.com/library/9a734d3d-78d2-1e27-63b3-2ad1074e16c1%28Office.15%29.aspx)|
|[PointsToEmus](http://msdn.microsoft.com/library/cb3f0bb9-fa0d-d967-9294-081a369c2c4e%28Office.15%29.aspx)|
|[PointsToInches](http://msdn.microsoft.com/library/58bfd9ce-dee7-0a14-8ec1-7e16a5e967d8%28Office.15%29.aspx)|
|[PointsToLines](http://msdn.microsoft.com/library/beab39fe-9458-6878-ae45-487a8b2271df%28Office.15%29.aspx)|
|[PointsToMillimeters](http://msdn.microsoft.com/library/eaa9154d-1a9b-81e7-58bc-3f7bf873ab97%28Office.15%29.aspx)|
|[PointsToPicas](http://msdn.microsoft.com/library/ff566bef-7032-70f7-7880-ff66cfeca88f%28Office.15%29.aspx)|
|[PointsToPixels](http://msdn.microsoft.com/library/9c67fcae-6c93-ddae-cbad-75356e5c5084%28Office.15%29.aspx)|
|[PointsToTwips](http://msdn.microsoft.com/library/ba928b83-f551-049e-5868-098a9837ee7b%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/db5a02ec-e553-6de1-0e2c-4a9a512e68fe%28Office.15%29.aspx)|
|[ShowWizardCatalog](http://msdn.microsoft.com/library/a8307ff9-a6c1-7655-8127-284f3781dae9%28Office.15%29.aspx)|
|[TwipsToPoints](http://msdn.microsoft.com/library/18e1c4da-1295-31a2-d66b-ab0df807b7a6%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveDocument](http://msdn.microsoft.com/library/c6293fa6-291c-d8ce-be54-f8a997b95d2e%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/125e2bb4-f922-ceef-9e3e-5dbe3aaff2a4%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/f3ed5997-b8ef-4729-4537-ae21424d2007%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/2abac248-bec5-876f-9ae5-88a59ce16b59%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/610f6300-0335-4fa1-7574-14afcf0e96e6%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/e0d4bb8e-5185-3d3c-fd80-c1e3c3902b2c%28Office.15%29.aspx)|
|[CaptionStyles](http://msdn.microsoft.com/library/d843db6a-b0e0-4ee0-a3ae-824c0c8391a9%28Office.15%29.aspx)|
|[ColorSchemes](http://msdn.microsoft.com/library/b991d8a2-d25d-839a-c14a-18cb6d126d33%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/b6f48f72-871a-6b7c-761c-9a9e0599acfa%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/21537c04-d406-6016-4f35-2f6ce6851db2%28Office.15%29.aspx)|
|[Documents](http://msdn.microsoft.com/library/dd48d68f-a6ae-b5c0-2a85-90abff1e6c5a%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/65d73a9d-be4c-d809-d10d-468181ef9eb0%28Office.15%29.aspx)|
|[InsertBarcodeVisible](http://msdn.microsoft.com/library/27b7f2aa-e7d7-5024-6c4a-75f2f275e924%28Office.15%29.aspx)|
|[InstalledPrinters](http://msdn.microsoft.com/library/e7cc1387-1ed8-dee8-a9f3-8c85eb1bea91%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/2fcfbec9-0c84-43d5-8c53-5b73bca17e3d%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/1abbf9ab-f7b4-1119-68c8-5c49d74a45b3%28Office.15%29.aspx)|
|[OfficeDataSourceObject](http://msdn.microsoft.com/library/d7262328-d5b6-6f55-d8c1-e6c072e29e3f%28Office.15%29.aspx)|
|[Options](http://msdn.microsoft.com/library/999f208a-02e6-49fb-c9a0-42aa97c5e37e%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/cab07b56-4c25-7309-5c06-bead2d5f691b%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/36ac9a9c-8235-aeba-c3d5-d39aef960cc5%28Office.15%29.aspx)|
|[PathSeparator](http://msdn.microsoft.com/library/f8c07ce4-d171-9c5b-60ac-d544bf65e620%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/a6606819-89d1-609d-62c3-c59159ff2ef7%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/aacd5ff6-dad1-af86-f4e0-af9012ae93f8%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/d265b4fb-1452-91a5-32fe-0cad54c8f29c%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/b4a542a7-cb54-476b-9ccf-004ce4b9ec47%28Office.15%29.aspx)|
|[ShowFollowUpCustom](http://msdn.microsoft.com/library/5853d057-f31b-d7e0-81fb-3e353e30709a%28Office.15%29.aspx)|
|[SnapToGuides](http://msdn.microsoft.com/library/09894c02-3193-cd14-ff55-45920e461af9%28Office.15%29.aspx)|
|[SnapToObjects](http://msdn.microsoft.com/library/84fcb808-bf3b-49f7-666e-915ac6b04a96%28Office.15%29.aspx)|
|[TemplateFolderPath](http://msdn.microsoft.com/library/e2256af9-9432-6205-864a-10bb7dec41c9%28Office.15%29.aspx)|
|[ValidateAddressVisible](http://msdn.microsoft.com/library/64d3732b-c549-c97b-511f-3122bb192ee5%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/ffec5bca-cd81-77c6-d80b-e629abfa6dec%28Office.15%29.aspx)|
|[WebOptions](http://msdn.microsoft.com/library/2e0c3435-a55a-4903-a0f8-9c347dec03b5%28Office.15%29.aspx)|
|[WizardCatalogVisible](http://msdn.microsoft.com/library/99323335-aabd-6799-b6aa-c5d95b88064f%28Office.15%29.aspx)|

## See also


#### Other resources


[Application Object Members](http://msdn.microsoft.com/library/aa4d515b-f779-b8b5-968a-8e5f7466fb56%28Office.15%29.aspx)
