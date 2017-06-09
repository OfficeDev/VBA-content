---
title: Document Object (Publisher)
keywords: vbapb10.chm553713663
f1_keywords:
- vbapb10.chm553713663
ms.prod: publisher
api_name:
- Publisher.Document
ms.assetid: 44f02255-ff5b-bcfe-900f-61c8fdf61ef3
ms.date: 06/08/2017
---


# Document Object (Publisher)

Represents a publication. 


## Example

Use the  **[ActiveDocument](http://msdn.microsoft.com/library/c6293fa6-291c-d8ce-be54-f8a997b95d2e%28Office.15%29.aspx)** property to refer to the current publication. This example adds a table to the first page of the active publication.


```
Sub NewTable() 
 With ActiveDocument.Pages(1).Shapes 
 .AddTable NumRows:=3, NumColumns:=3, Left:=72, Top:=300, _ 
 Width:=488, Height:=36 
 With .Item(1).Table.Rows(1) 
 .Cells(1).TextRange.Text = "Column1" 
 .Cells(2).TextRange.Text = "Column2" 
 .Cells(3).TextRange.Text = "Column3" 
 End With 
 End With 
End Sub
```

You can also write the above routine by using a reference to the  **ThisDocument** module. This example uses a **ThisDocument** reference instead of **ActiveDocument**.




```
Sub PrintPublication() 
 With ThisDocument.Pages(1).Shapes 
 .AddTable NumRows:=3, NumColumns:=3, Left:=72, Top:=300, _ 
 Width:=488, Height:=36 
 With .Item(1).Table.Rows(1) 
 .Cells(1).TextRange.Text = "Column1" 
 .Cells(2).TextRange.Text = "Column2" 
 .Cells(3).TextRange.Text = "Column3" 
 End With 
 End With 
End Sub
```


## Events



|**Name**|
|:-----|
|[BeforeClose](http://msdn.microsoft.com/library/d40e36b6-fea7-a9d5-0c88-55197983b888%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/43108d1d-d101-8a07-943e-c9b8dbadcbfd%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/c00db13d-1c03-2536-8923-bd7d9393fee2%28Office.15%29.aspx)|
|[ShapesAdded](http://msdn.microsoft.com/library/f6573f7c-56fa-1efa-9dba-39cde3859cc0%28Office.15%29.aspx)|
|[ShapesRemoved](http://msdn.microsoft.com/library/e2a67359-5673-2c72-e1fc-e3e3a3b564f9%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/9789e469-dc84-a0b7-ffe0-405d4e7ad861%28Office.15%29.aspx)|
|[WizardAfterChange](http://msdn.microsoft.com/library/c4ec0950-3a58-1f29-b35f-35db9d87f330%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[BeginCustomUndoAction](http://msdn.microsoft.com/library/316f443e-6782-594b-b955-f5ab60140f6a%28Office.15%29.aspx)|
|[ChangeDocument](http://msdn.microsoft.com/library/c6defa92-99fb-973b-6bb2-e3c2a1b0a4f3%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/b4b21484-1858-b7b3-291f-18ef8cab8ba7%28Office.15%29.aspx)|
|[ConvertPublicationType](http://msdn.microsoft.com/library/e4bfe349-a22f-6017-ac9d-49f67e1f6dd2%28Office.15%29.aspx)|
|[EndCustomUndoAction](http://msdn.microsoft.com/library/5b703366-8d0e-1bbc-3320-a2fea99468c3%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/8bb5b64f-57b2-cf87-344c-be1e2741a59c%28Office.15%29.aspx)|
|[FindShapeByWizardTag](http://msdn.microsoft.com/library/c6db9ba7-15b0-e8f0-1ed2-08b6e978c948%28Office.15%29.aspx)|
|[FindShapesByTag](http://msdn.microsoft.com/library/405a0f39-5892-23da-904a-5188a4340b00%28Office.15%29.aspx)|
|[PrintOutEx](http://msdn.microsoft.com/library/f11b6f8b-08a0-28f6-5930-47d684585bef%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/4b76aeaa-77f7-5f22-ff80-77479b0f0702%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/89eae461-d1c2-b3ca-58b7-9528df8801d8%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/ba8b85d7-8ca9-dcf5-12b4-4cabced743e6%28Office.15%29.aspx)|
|[SetBusinessInformation](http://msdn.microsoft.com/library/8549f75f-2fb6-6ac6-ecaf-54a0a9b22dc7%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/8cfd09a0-8a0d-2870-f833-a35ff1fc21b4%28Office.15%29.aspx)|
|[UndoClear](http://msdn.microsoft.com/library/63e9bb00-950f-3e30-3897-434362b9efbf%28Office.15%29.aspx)|
|[UpdateOLEObjects](http://msdn.microsoft.com/library/2c07e755-6f5c-5fd8-091c-fbe3bfae6692%28Office.15%29.aspx)|
|[WebPagePreview](http://msdn.microsoft.com/library/44083fae-d21d-9cd3-3553-a4d4346141f5%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveView](http://msdn.microsoft.com/library/1448c8c6-30e5-2e2a-f124-ebf544d8f297%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/0d00a8fa-aef2-43df-3c54-0cca804b7eee%28Office.15%29.aspx)|
|[AdvancedPrintOptions](http://msdn.microsoft.com/library/33c075e0-f813-9bb4-e199-96e5e9ed4ba8%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/eb401e80-3101-a19f-dc62-5386d123ac7d%28Office.15%29.aspx)|
|[AvailableBuildingBlocks](http://msdn.microsoft.com/library/dab447d9-f044-4a40-8876-a96f233b8d2e%28Office.15%29.aspx)|
|[BorderArts](http://msdn.microsoft.com/library/5639ffce-f711-71b6-78f8-2de63fe50a3c%28Office.15%29.aspx)|
|[ColorScheme](http://msdn.microsoft.com/library/b7748b48-eff3-bdf0-e6ce-a9a2e788d0f7%28Office.15%29.aspx)|
|[DefaultTabStop](http://msdn.microsoft.com/library/245ff7a3-9828-5220-b692-2ce6effb9eb6%28Office.15%29.aspx)|
|[DocumentDirection](http://msdn.microsoft.com/library/b28961ad-7adc-3920-0e67-88bb53310d9b%28Office.15%29.aspx)|
|[EnvelopeVisible](http://msdn.microsoft.com/library/65423c1f-e61b-3c83-4bff-ddd278d97238%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/e9b31937-4504-79b5-5913-b2ef0a23f2a7%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/137e4310-8431-ed2a-503a-c225378a9a74%28Office.15%29.aspx)|
|[IsDataSourceConnected](http://msdn.microsoft.com/library/b62422ab-12f7-1151-d8d1-1cb32de18160%28Office.15%29.aspx)|
|[IsWizard](http://msdn.microsoft.com/library/61ee1a16-eccb-908f-2b34-eee03175c37e%28Office.15%29.aspx)|
|[LayoutGuides](http://msdn.microsoft.com/library/0c45366d-6b7a-7cf3-a566-bb945ff32ba4%28Office.15%29.aspx)|
|[MailEnvelope](http://msdn.microsoft.com/library/3c4c734a-6725-5f6e-ed0a-5b19e4e642bd%28Office.15%29.aspx)|
|[MailMerge](http://msdn.microsoft.com/library/15b1a8aa-3472-c67d-1d99-92617b05c157%28Office.15%29.aspx)|
|[MasterPages](http://msdn.microsoft.com/library/26e5342b-94f0-4fd5-2743-92cfd2d43a01%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/fcf86fcc-a3aa-b4c6-1ecc-202972ac558b%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/2bb3e529-a459-b37c-c9ae-4cc059954a63%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/1dac39f0-2507-a85b-8c71-cd1980022fb3%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/d9081ba8-f0ae-a68a-a5a0-56c4a7caf422%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/01926d63-e59e-5aad-3cb9-143166d253a5%28Office.15%29.aspx)|
|[PrintPageBackgrounds](http://msdn.microsoft.com/library/6d1d6e6a-fd66-2afa-2172-4a6552d5cce4%28Office.15%29.aspx)|
|[PrintStyle](http://msdn.microsoft.com/library/ac9c8bc0-3c03-d094-fdda-1f2f5966f717%28Office.15%29.aspx)|
|[PublicationType](http://msdn.microsoft.com/library/264c2769-2452-0009-4853-84a6a426db38%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/9ee6488d-3070-e784-e772-78dace2c1284%28Office.15%29.aspx)|
|[RedoActionsAvailable](http://msdn.microsoft.com/library/9af11772-e807-730a-89a0-da06e979f834%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/bbc1aee1-90ca-966e-c17c-579064318cd1%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/d1f4357a-103c-2227-d1bd-50706e1f241c%28Office.15%29.aspx)|
|[SaveFormat](http://msdn.microsoft.com/library/545f0411-899f-ffe3-e844-8c2922a357f0%28Office.15%29.aspx)|
|[ScratchArea](http://msdn.microsoft.com/library/782d9b7f-b620-60f0-c21d-04f588c37cc6%28Office.15%29.aspx)|
|[Sections](http://msdn.microsoft.com/library/9e425836-1d62-99ef-2984-b61f3a3cf831%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/b1098cdb-8fb7-0906-b193-6dc572ac2993%28Office.15%29.aspx)|
|[Stories](http://msdn.microsoft.com/library/4ffc7d20-eb11-942e-e28a-81c2caa19a50%28Office.15%29.aspx)|
|[SurplusShapes](http://msdn.microsoft.com/library/8c1c5fee-bea0-1660-a4a5-b465879d6ec9%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/d8baaf50-86ad-1997-c1b3-e54a77a3ee5b%28Office.15%29.aspx)|
|[TextStyles](http://msdn.microsoft.com/library/a628e5c1-aed7-dd70-81fa-d9fb54afb527%28Office.15%29.aspx)|
|[UndoActionsAvailable](http://msdn.microsoft.com/library/1dd20295-3987-c36d-ccc1-9e18a7887f33%28Office.15%29.aspx)|
|[ViewBoundaries](http://msdn.microsoft.com/library/6e390607-a3f4-f938-4a3f-75d8a993cf2a%28Office.15%29.aspx)|
|[ViewGuides](http://msdn.microsoft.com/library/a0533bc6-8565-eb4f-67e3-b438d4460e80%28Office.15%29.aspx)|
|[ViewHorizontalBaseLineGuides](http://msdn.microsoft.com/library/e5471313-38e0-9454-04af-4c85d976b312%28Office.15%29.aspx)|
|[ViewTwoPageSpread](http://msdn.microsoft.com/library/b5e851ff-d5fc-a98d-02b3-7e14c1b957dc%28Office.15%29.aspx)|
|[ViewVerticalBaseLineGuides](http://msdn.microsoft.com/library/711335ab-237b-65a2-534a-7635cfba474e%28Office.15%29.aspx)|
|[WebNavigationBarSets](http://msdn.microsoft.com/library/4193dbce-a2e3-2587-5282-43b4c3cec921%28Office.15%29.aspx)|
|[Wizard](http://msdn.microsoft.com/library/26603c80-2b03-9889-27d7-623e71f84b74%28Office.15%29.aspx)|

