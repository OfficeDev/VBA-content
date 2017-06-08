---
title: Document Object (Visio)
keywords: vis_sdr.chm10080
f1_keywords:
- vis_sdr.chm10080
ms.prod: visio
api_name:
- Visio.Document
ms.assetid: 21640062-13a2-a2b2-7c61-7e707671207c
ms.date: 06/08/2017
---


# Document Object (Visio)

Represents a drawing file (.vsd or .vdx), stencil file (.vss or .vsx), or template file (.vst or .vtx) that is open in an instance of Microsoft Visio. A  **Document** object is a member of the **Documents** collection of an **Application** object.


## Remarks

The default property of a  **Document** object is **Name**.

Use the  **Open** method of a **Documents** collection to open an existing document.

Use the  **Add** method of a **Documents** collection to create a new document.

Use the  **ActiveDocument** property of an **Application** object to retrieve the active document in an instance.

Use the  **Pages**, **Masters**, and **Styles** properties of a **Document** object to retrieve **Page**, **Master**, and **Style** objects, respectively.


 **Note**  

Use the  **CustomMenus** or **CustomToolbars** properties of a **Document** object to access the custom menus or toolbars.


 **Note**   The Microsoft Visual Basic for Applications (VBA) project of every Visio document also has a class module called **ThisDocument**. When you reference the **ThisDocument** module from code in a VBA project, it returns a reference to the project's **Document** object. For example, the code in a document's project can display the name of the project's document in a **message** box with this statement:




```
    MsgBox ThisDocument.Name
```

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this object maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVDocument**
    

## Events



|**Name**|
|:-----|
|[AfterDocumentMerge](http://msdn.microsoft.com/library/50658da5-592a-4d16-908f-c6abe3050f09%28Office.15%29.aspx)|
|[AfterRemoveHiddenInformation](http://msdn.microsoft.com/library/d407a676-1917-f24f-7651-ad2f05872b91%28Office.15%29.aspx)|
|[BeforeDataRecordsetDelete](http://msdn.microsoft.com/library/6d9d8570-bdfd-0762-4531-116589203bed%28Office.15%29.aspx)|
|[BeforeDocumentClose](http://msdn.microsoft.com/library/e35f9593-f5ee-f84b-95e6-f23a899c0d6d%28Office.15%29.aspx)|
|[BeforeDocumentSave](http://msdn.microsoft.com/library/03f8954d-40d7-fb64-8c83-cc8f6ca66653%28Office.15%29.aspx)|
|[BeforeDocumentSaveAs](http://msdn.microsoft.com/library/6802441e-5020-8d5c-f637-3654df71cba0%28Office.15%29.aspx)|
|[BeforeMasterDelete](http://msdn.microsoft.com/library/5f482099-7b42-de36-6e51-34ff463a49ed%28Office.15%29.aspx)|
|[BeforePageDelete](http://msdn.microsoft.com/library/dd41d679-d6f7-524f-c714-bea38ae1a0b4%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/25fccddf-efbb-8041-087a-2c3e3b5cc12c%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/e97cb920-7830-0e84-b299-cc305fbb4feb%28Office.15%29.aspx)|
|[BeforeStyleDelete](http://msdn.microsoft.com/library/dd6b89f8-0b4c-1ca2-aae8-e9781f4ef50f%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/3a782db3-2df8-287b-dd42-dce73b24b7cb%28Office.15%29.aspx)|
|[DataRecordsetAdded](http://msdn.microsoft.com/library/3ddb399d-0b28-9ec7-4059-f8d3011a98c0%28Office.15%29.aspx)|
|[DesignModeEntered](http://msdn.microsoft.com/library/c8fc31b5-8770-f068-d469-aeb110214824%28Office.15%29.aspx)|
|[DocumentChanged](http://msdn.microsoft.com/library/3a7fd39e-f944-1c41-a5d3-130e795836bf%28Office.15%29.aspx)|
|[DocumentCloseCanceled](http://msdn.microsoft.com/library/f553b8d5-0531-4bc6-d27d-315193b76e0b%28Office.15%29.aspx)|
|[DocumentCreated](http://msdn.microsoft.com/library/5d5c0c99-fce1-13fb-a2e1-98f829784ee6%28Office.15%29.aspx)|
|[DocumentOpened](http://msdn.microsoft.com/library/32e1d16e-1906-9477-bdb7-e72833a055f2%28Office.15%29.aspx)|
|[DocumentSaved](http://msdn.microsoft.com/library/48e513a1-4382-eb3c-4838-ad2f85483f51%28Office.15%29.aspx)|
|[DocumentSavedAs](http://msdn.microsoft.com/library/36714188-964b-880b-9504-62a6a50179f1%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/135d8176-2c26-12aa-5bff-0df205e0640f%28Office.15%29.aspx)|
|[MasterAdded](http://msdn.microsoft.com/library/5637df50-5174-03d4-a07f-cc7aeb92d0fa%28Office.15%29.aspx)|
|[MasterChanged](http://msdn.microsoft.com/library/59fe2ee8-03ee-83b9-d86c-a67d68c7a363%28Office.15%29.aspx)|
|[MasterDeleteCanceled](http://msdn.microsoft.com/library/e2d82547-46a9-7994-e317-78be658208c6%28Office.15%29.aspx)|
|[PageAdded](http://msdn.microsoft.com/library/3a49fcb4-fa41-e13e-ea2c-beb87aff3e40%28Office.15%29.aspx)|
|[PageChanged](http://msdn.microsoft.com/library/ab5b9492-60d5-35c2-642c-14e588e79f7d%28Office.15%29.aspx)|
|[PageDeleteCanceled](http://msdn.microsoft.com/library/f4a81afb-42b5-723b-b5e6-6505e12f538f%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/1199e5ac-26b5-c5ca-106f-1ff4b833b933%28Office.15%29.aspx)|
|[QueryCancelDocumentClose](http://msdn.microsoft.com/library/e00d4708-24dd-3a35-c986-54464a028a6b%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/0fb4f654-f501-32d7-d94d-5240cfc82eb4%28Office.15%29.aspx)|
|[QueryCancelMasterDelete](http://msdn.microsoft.com/library/b363d3d7-e3ca-2cd2-bd29-b224de7cadc8%28Office.15%29.aspx)|
|[QueryCancelPageDelete](http://msdn.microsoft.com/library/d4f59122-5e03-72f8-5a9d-23e629a658a4%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/6b784ad0-a8fb-dd07-9e87-abaa1509af1b%28Office.15%29.aspx)|
|[QueryCancelStyleDelete](http://msdn.microsoft.com/library/07417cc7-f535-4217-8a4d-09cd7e5d5b84%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/e25505a9-a2ae-dc68-8bf6-ac4252c7f5e6%28Office.15%29.aspx)|
|[RuleSetValidated](http://msdn.microsoft.com/library/682b8f48-4ebe-ce53-f816-3d82a4ae0034%28Office.15%29.aspx)|
|[RunModeEntered](http://msdn.microsoft.com/library/8e582dd1-b2c5-72e5-b144-510726d35a18%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/43638a89-c047-33fb-ea05-13d217979102%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/d80b6ee3-8b5f-9c34-e8db-8443146b4728%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/05a38afb-520d-06a7-c62e-58aa4ae653e1%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/8c511847-f5e1-d5af-e375-c9f4153b7515%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/0397a034-6b79-c760-9bbf-759f62109cef%28Office.15%29.aspx)|
|[StyleAdded](http://msdn.microsoft.com/library/e6bed9a7-e448-061d-3547-a383697ffdc3%28Office.15%29.aspx)|
|[StyleChanged](http://msdn.microsoft.com/library/1e07a517-4c3f-12a1-896e-0b9262b5736e%28Office.15%29.aspx)|
|[StyleDeleteCanceled](http://msdn.microsoft.com/library/e5484540-cf9d-0cbf-acb7-0ab9dad8b7c2%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/e7ba2c59-b43c-e89f-7921-0a2e624bcad5%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddUndoUnit](http://msdn.microsoft.com/library/3b9d903a-8854-fa64-a9c5-85ac71d58f54%28Office.15%29.aspx)|
|[BeginUndoScope](http://msdn.microsoft.com/library/4e0c99a3-3ac6-54f8-3e43-1c79224e09e1%28Office.15%29.aspx)|
|[CanCheckIn](http://msdn.microsoft.com/library/99922339-631b-f60e-1d07-3ae9df134cf7%28Office.15%29.aspx)|
|[CanUndoCheckOut](http://msdn.microsoft.com/library/aa271635-73ef-b681-364c-49d515fd54cb%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/9b75d468-24bc-e205-cafa-6e585f469e38%28Office.15%29.aspx)|
|[Clean](http://msdn.microsoft.com/library/5fd5c6a6-1914-b29d-c0ae-0e5374d13a8e%28Office.15%29.aspx)|
|[ClearCustomMenus](http://msdn.microsoft.com/library/5be16274-151b-e139-8607-76fdb05a4235%28Office.15%29.aspx)|
|[ClearCustomToolbars](http://msdn.microsoft.com/library/823877b1-ee82-f87e-d68f-d8c6010457cc%28Office.15%29.aspx)|
|[ClearGestureFormatSheet](http://msdn.microsoft.com/library/46f411b0-b822-cc7c-62e3-0b756d211d5d%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/913572fd-cacb-8d06-0e5f-3bd2e98d6d13%28Office.15%29.aspx)|
|[CopyPreviewPicture](http://msdn.microsoft.com/library/a0d5799e-700c-6dd6-ce91-08c8eecda79f%28Office.15%29.aspx)|
|[DeleteSolutionXMLElement](http://msdn.microsoft.com/library/2f00680e-56b1-c99b-2739-9d331965f802%28Office.15%29.aspx)|
|[Drop](http://msdn.microsoft.com/library/1e6b2d14-71c2-4adc-a9d7-ec123b2b7f31%28Office.15%29.aspx)|
|[EndUndoScope](http://msdn.microsoft.com/library/3a884984-7e45-8afd-3291-b706c8edab25%28Office.15%29.aspx)|
|[ExecuteLine](http://msdn.microsoft.com/library/0443c879-b569-c35b-e28c-77d0bf4b23ba%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/70b83f7e-b7f8-7b8f-d9d7-7f7b30f3b45d%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/555e607d-7e4a-d3c8-9c78-1733b112200c%28Office.15%29.aspx)|
|[GetThemeNames](http://msdn.microsoft.com/library/63477332-5db2-40ff-6918-7ab20a9f0fd0%28Office.15%29.aspx)|
|[GetThemeNamesU](http://msdn.microsoft.com/library/7a7280ae-10c9-9bc7-c121-29791e4df557%28Office.15%29.aspx)|
|[OpenStencilWindow](http://msdn.microsoft.com/library/70c3720b-b88d-4859-684b-5c7ae9c868ea%28Office.15%29.aspx)|
|[ParseLine](http://msdn.microsoft.com/library/46603de4-afa0-7903-f585-0a1aaa5c74c7%28Office.15%29.aspx)|
|[Print](http://msdn.microsoft.com/library/b7860f50-8027-cd2c-38db-0d7b9f17546a%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/c13f7640-7439-0c73-cde5-223b8b4549d3%28Office.15%29.aspx)|
|[PurgeUndo](http://msdn.microsoft.com/library/04556300-8787-5a04-040c-476d864f682e%28Office.15%29.aspx)|
|[RemoveHiddenInformation](http://msdn.microsoft.com/library/cc097f8b-5e74-9b44-4ba9-19537169c88b%28Office.15%29.aspx)|
|[RenameCurrentScope](http://msdn.microsoft.com/library/08aff947-e876-29b8-e910-e2a3b42e5d0e%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/5a9f104c-4893-c401-0093-bc860adf9a4b%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/308e92b1-de61-9ce3-19be-b7f9126247a0%28Office.15%29.aspx)|
|[SaveAsEx](http://msdn.microsoft.com/library/c0bef38d-1849-67ab-606f-8178de46c7c6%28Office.15%29.aspx)|
|[SetCustomMenus](http://msdn.microsoft.com/library/05d373a4-3aec-a427-57aa-94fc3ac10161%28Office.15%29.aspx)|
|[SetCustomToolbars](http://msdn.microsoft.com/library/fddae53c-0519-90ef-0023-ee3896e86757%28Office.15%29.aspx)|
|[UndoCheckOut](http://msdn.microsoft.com/library/7b6a67ae-2acd-217f-42e0-f8aced97ac96%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AlternateNames](http://msdn.microsoft.com/library/2d0a3f45-e9b4-385b-23c9-2a0a70375202%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/8643d912-21b2-18b4-e0fe-cc6e9db6ae58%28Office.15%29.aspx)|
|[AutoRecover](http://msdn.microsoft.com/library/23b09910-35a8-93bc-71f0-4618b1c48523%28Office.15%29.aspx)|
|[BottomMargin](http://msdn.microsoft.com/library/5fd185a5-ecc9-000e-f5b0-fa309d52847a%28Office.15%29.aspx)|
|[BuildNumberCreated](http://msdn.microsoft.com/library/a7fb5bad-ca87-820a-be93-458ad280b9d0%28Office.15%29.aspx)|
|[BuildNumberEdited](http://msdn.microsoft.com/library/91d39eb1-f416-6167-96af-53c5cf0ee35c%28Office.15%29.aspx)|
|[Category](http://msdn.microsoft.com/library/da312b56-6232-9077-e47b-47144aa603c5%28Office.15%29.aspx)|
|[ClassID](http://msdn.microsoft.com/library/668fec9a-eadf-a496-5db3-b91e30237c11%28Office.15%29.aspx)|
|[Colors](http://msdn.microsoft.com/library/e7ed0aa2-c365-bbf7-e06c-5df34094dd9a%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/15a322ad-70eb-1487-701d-76e2fde73309%28Office.15%29.aspx)|
|[Company](http://msdn.microsoft.com/library/b55e23dc-3b58-c062-1738-74d2f50fa39d%28Office.15%29.aspx)|
|[CompatibilityMode](http://msdn.microsoft.com/library/98fc00d3-5d2b-218e-9828-b5581ee7313d%28Office.15%29.aspx)|
|[Container](http://msdn.microsoft.com/library/a5b2c90e-f9e0-cc09-8388-566729c1c4eb%28Office.15%29.aspx)|
|[ContainsWorkspaceEx](http://msdn.microsoft.com/library/074d4b49-cb26-5d11-d710-7d27f2f4dd01%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/c1dea222-796c-1231-bb9b-9d258450b142%28Office.15%29.aspx)|
|[CustomMenus](http://msdn.microsoft.com/library/f7d3ec25-a62e-ffe3-affe-c80ed46f58a6%28Office.15%29.aspx)|
|[CustomMenusFile](http://msdn.microsoft.com/library/a35dea4c-be19-8951-516b-bc8de4345d78%28Office.15%29.aspx)|
|[CustomToolbars](http://msdn.microsoft.com/library/def64862-5298-bc3a-0509-84216725d7da%28Office.15%29.aspx)|
|[CustomToolbarsFile](http://msdn.microsoft.com/library/1385e027-0cc9-4f3b-a044-ff5731325b25%28Office.15%29.aspx)|
|[CustomUI](http://msdn.microsoft.com/library/dff5841d-f2cc-c8fd-1b30-ca0145f5c04c%28Office.15%29.aspx)|
|[DataRecordsets](http://msdn.microsoft.com/library/d15182ba-27e7-ab0e-6ac0-c23515848032%28Office.15%29.aspx)|
|[DefaultFillStyle](http://msdn.microsoft.com/library/c013a054-99ef-2bc1-196d-f3877289a278%28Office.15%29.aspx)|
|[DefaultGuideStyle](http://msdn.microsoft.com/library/d739d6ca-01c4-d99b-df32-d2589f015fb7%28Office.15%29.aspx)|
|[DefaultLineStyle](http://msdn.microsoft.com/library/6a1d7752-25c9-638e-9e10-02928849a8db%28Office.15%29.aspx)|
|[DefaultSavePath](http://msdn.microsoft.com/library/13d159c8-294b-aa3f-466d-092f3ef0b93c%28Office.15%29.aspx)|
|[DefaultStyle](http://msdn.microsoft.com/library/e8fb078f-72cd-b4ae-1685-c0c02a265d3e%28Office.15%29.aspx)|
|[DefaultTextStyle](http://msdn.microsoft.com/library/cae34239-14af-92c3-a498-8ac7f51e1fa0%28Office.15%29.aspx)|
|[Description](http://msdn.microsoft.com/library/530adbe3-5285-6aa5-32e6-88d2bc1d8ebf%28Office.15%29.aspx)|
|[DiagramServicesEnabled](http://msdn.microsoft.com/library/1a492029-31c8-85bb-0843-31c0a1200055%28Office.15%29.aspx)|
|[DocumentSheet](http://msdn.microsoft.com/library/914bf120-2f7c-6a2e-0f8a-a6b05252f49b%28Office.15%29.aspx)|
|[DynamicGridEnabled](http://msdn.microsoft.com/library/07c49f2e-7d37-681c-7c49-b07e1d99da0c%28Office.15%29.aspx)|
|[EditorCount](http://msdn.microsoft.com/library/36e90125-e217-4818-ad9c-58a52c88dd8a%28Office.15%29.aspx)|
|[EmailRoutingData](http://msdn.microsoft.com/library/28dfec3c-d929-efe4-bbac-2816e6b70f0e%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/a23efd7e-6591-9663-6c90-6f006e2602db%28Office.15%29.aspx)|
|[Fonts](http://msdn.microsoft.com/library/061ecb2f-b36f-3bf2-0da8-b95f7cc52031%28Office.15%29.aspx)|
|[FooterCenter](http://msdn.microsoft.com/library/7abdcd6c-39ed-ad05-bfef-cbd979f3a8d6%28Office.15%29.aspx)|
|[FooterLeft](http://msdn.microsoft.com/library/e832c09d-3ddb-4351-43ad-e1c5633b7bc9%28Office.15%29.aspx)|
|[FooterMargin](http://msdn.microsoft.com/library/f35ea698-bfff-7c46-a4ee-8faf9cc4ac27%28Office.15%29.aspx)|
|[FooterRight](http://msdn.microsoft.com/library/17db938c-6b1b-6cd1-7f4e-65aca275f30b%28Office.15%29.aspx)|
|[FullBuildNumberCreated](http://msdn.microsoft.com/library/3520525a-4c76-3583-49a6-015f2fb90366%28Office.15%29.aspx)|
|[FullBuildNumberEdited](http://msdn.microsoft.com/library/43a6ff61-2ab8-4e89-0e06-bd2ba6ec0f02%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/9f6d15ab-9913-57f4-a0ee-57618d5b1b0f%28Office.15%29.aspx)|
|[GestureFormatSheet](http://msdn.microsoft.com/library/26d3c27f-31ff-198c-5b40-8dc8b30b6362%28Office.15%29.aspx)|
|[GlueEnabled](http://msdn.microsoft.com/library/fdcda6ec-2390-95e7-d5d2-2d1048991d2e%28Office.15%29.aspx)|
|[GlueSettings](http://msdn.microsoft.com/library/192fb40f-d244-48e9-59ad-4439385bf3e5%28Office.15%29.aspx)|
|[HeaderCenter](http://msdn.microsoft.com/library/8695883a-8b00-eef4-aecd-81ad47581a82%28Office.15%29.aspx)|
|[HeaderFooterColor](http://msdn.microsoft.com/library/d1f3887f-d6b5-feb5-b119-ddf3d9eb3542%28Office.15%29.aspx)|
|[HeaderFooterFont](http://msdn.microsoft.com/library/cd4b1f35-c3a2-d48c-fc0d-37f9626ecdab%28Office.15%29.aspx)|
|[HeaderLeft](http://msdn.microsoft.com/library/f19dede9-e28b-8fc4-bbbd-82b0e72d48c9%28Office.15%29.aspx)|
|[HeaderMargin](http://msdn.microsoft.com/library/7d2c137d-6b75-9747-5a6a-5e5d99156d45%28Office.15%29.aspx)|
|[HeaderRight](http://msdn.microsoft.com/library/3d702cb7-9b70-5f00-c2ea-b619cbfed37f%28Office.15%29.aspx)|
|[HyperlinkBase](http://msdn.microsoft.com/library/cde4801e-269d-b6d3-aee1-d95b2e36bfd2%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/186eb9ff-eed4-b554-4885-aa0e05e88ce4%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/f72e68b9-c249-b4df-14ae-669509100546%28Office.15%29.aspx)|
|[InPlace](http://msdn.microsoft.com/library/8bd0c927-3314-5228-11d6-291a54fd7cdb%28Office.15%29.aspx)|
|[Keywords](http://msdn.microsoft.com/library/c7717a93-c64f-8363-69a7-7e9ed40865dc%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/76f995fd-8b4d-7292-50c1-8dcb6448c2ec%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/9f880830-8b63-2a34-2a02-fd6b6a225c7a%28Office.15%29.aspx)|
|[MacrosEnabled](http://msdn.microsoft.com/library/361b7bad-55f9-2d4b-4de3-8a12da48d59e%28Office.15%29.aspx)|
|[Manager](http://msdn.microsoft.com/library/6aa5bcfc-55b5-88ce-a9a8-d0f6a73ee69f%28Office.15%29.aspx)|
|[Masters](http://msdn.microsoft.com/library/b139014c-6d7c-ba76-8366-bcacecc5c639%28Office.15%29.aspx)|
|[MasterShortcuts](http://msdn.microsoft.com/library/7d156dfe-ac70-355a-5927-eb7ebb28bb21%28Office.15%29.aspx)|
|[Mode](http://msdn.microsoft.com/library/40ebcc64-43dc-79f4-2802-9cd9dba633ab%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/91b1a838-2f0c-56be-4d23-ab9f5a157964%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/4d981d9d-67ba-81d2-d1c0-34810b24af92%28Office.15%29.aspx)|
|[OLEObjects](http://msdn.microsoft.com/library/3cb58d69-2287-2dbc-a6fb-f8a1ec9cf854%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/db81b42f-dfd7-c4dc-a520-b1927cd1e737%28Office.15%29.aspx)|
|[PaperHeight](http://msdn.microsoft.com/library/305356e8-69d6-bae3-5136-d931fcf967b5%28Office.15%29.aspx)|
|[PaperSize](http://msdn.microsoft.com/library/a51b3d26-e79e-d572-055f-fc1c4a94096e%28Office.15%29.aspx)|
|[PaperWidth](http://msdn.microsoft.com/library/e43d7d44-31ad-24e3-79e4-6005cbd65612%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/50c20d69-3909-9383-1d2c-d1744a96e751%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/eaa00c97-f2ae-32c6-fe72-32c866d2476c%28Office.15%29.aspx)|
|[PreviewPicture](http://msdn.microsoft.com/library/4354f66b-6f0b-1511-3c77-fc7cd58f539e%28Office.15%29.aspx)|
|[PrintCenteredH](http://msdn.microsoft.com/library/f91d63cf-e447-1e1a-2c45-c2a48d0ab4dc%28Office.15%29.aspx)|
|[PrintCenteredV](http://msdn.microsoft.com/library/e60866c2-e6cf-3d42-1443-0a4cbedb5609%28Office.15%29.aspx)|
|[Printer](http://msdn.microsoft.com/library/cb710f0e-a284-c81a-c45d-5bf66d508743%28Office.15%29.aspx)|
|[PrintFitOnPages](http://msdn.microsoft.com/library/d129ad36-0728-b3b5-60b5-3ba52e102cc7%28Office.15%29.aspx)|
|[PrintLandscape](http://msdn.microsoft.com/library/4279a23b-2de8-3fbe-77b1-4b7bdd8db374%28Office.15%29.aspx)|
|[PrintPagesAcross](http://msdn.microsoft.com/library/43c09ce5-fcc9-d91c-3108-5e2dcb658f74%28Office.15%29.aspx)|
|[PrintPagesDown](http://msdn.microsoft.com/library/eacf72d7-f784-7a2b-0579-8af7991430ea%28Office.15%29.aspx)|
|[PrintScale](http://msdn.microsoft.com/library/d352b695-1e94-888d-70a0-9189678992e6%28Office.15%29.aspx)|
|[ProgID](http://msdn.microsoft.com/library/a3ae063b-8054-a5c7-4afd-2dac64ea6537%28Office.15%29.aspx)|
|[Protection](http://msdn.microsoft.com/library/f80cd284-e0e3-0663-c505-88311ffc9d3b%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/0645ee7b-7d51-a89d-b2ec-987037397eb8%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/b05eb59e-9906-10f9-8819-60f8f0f1d4f5%28Office.15%29.aspx)|
|[RightMargin](http://msdn.microsoft.com/library/ee2fc9f4-92a6-a787-7fa0-cd43da52fadb%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/de3141f6-eda9-a62b-847c-e946966fae6b%28Office.15%29.aspx)|
|[SavePreviewMode](http://msdn.microsoft.com/library/e40f2b06-c9fd-3133-73c9-306f46f21e55%28Office.15%29.aspx)|
|[ServerPublishOptions](http://msdn.microsoft.com/library/95d7b668-3a72-a15c-550d-18ef02d0309f%28Office.15%29.aspx)|
|[SharedWorkspace](http://msdn.microsoft.com/library/100d635c-2b2a-4ba3-0490-bc4a4c4efb8c%28Office.15%29.aspx)|
|[SnapAngles](http://msdn.microsoft.com/library/b2a85580-3308-6bda-dd00-7449f6d87c8b%28Office.15%29.aspx)|
|[SnapEnabled](http://msdn.microsoft.com/library/d2f7b068-b8a8-21d1-9a34-82d693fe2cad%28Office.15%29.aspx)|
|[SnapExtensions](http://msdn.microsoft.com/library/8b5aad7a-335a-dc8c-aa58-42947ffdc53e%28Office.15%29.aspx)|
|[SnapSettings](http://msdn.microsoft.com/library/c3ced586-d9c7-01bd-6b32-99fedda3c2b8%28Office.15%29.aspx)|
|[SolutionXMLElement](http://msdn.microsoft.com/library/44e9daa6-96dc-3041-ed50-dd4670298b6a%28Office.15%29.aspx)|
|[SolutionXMLElementCount](http://msdn.microsoft.com/library/da72e807-749b-fe05-578b-89289bce970d%28Office.15%29.aspx)|
|[SolutionXMLElementExists](http://msdn.microsoft.com/library/d4a0bd9b-a3ea-de0a-5c33-ccad4d4398eb%28Office.15%29.aspx)|
|[SolutionXMLElementName](http://msdn.microsoft.com/library/460993bc-090c-00ad-805f-ae4af832ceba%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/4121b945-ab6c-ce15-9441-78e031907004%28Office.15%29.aspx)|
|[Styles](http://msdn.microsoft.com/library/41434c49-3306-78b5-2126-0320fc05825a%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/b954ca88-c7f7-0c1f-ed30-8ea3eb3bc0e3%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/1e5ef6da-a665-024f-5e35-e8518f4d1054%28Office.15%29.aspx)|
|[Template](http://msdn.microsoft.com/library/c9e579d7-4448-4dc7-0130-1b38d41cbf1a%28Office.15%29.aspx)|
|[Time](http://msdn.microsoft.com/library/04d7d5d9-9e4f-c64a-faa9-ac521807b44f%28Office.15%29.aspx)|
|[TimeCreated](http://msdn.microsoft.com/library/efb0fdc6-c4ff-78a5-08bb-7a4367cedc43%28Office.15%29.aspx)|
|[TimeEdited](http://msdn.microsoft.com/library/2c4efd8a-ae6a-69b0-5033-b456f84f5acf%28Office.15%29.aspx)|
|[TimePrinted](http://msdn.microsoft.com/library/f5dd01f0-42dc-ab0d-4cd2-c85da6181ea0%28Office.15%29.aspx)|
|[TimeSaved](http://msdn.microsoft.com/library/801c7940-b838-15ae-cee8-e07ca5ae78ea%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/9a3b9e5f-2515-dda4-d757-0a0f375dfd00%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/ed8d16c2-f80d-d444-28a4-d9f0db4ab6d3%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/87def9ff-e9f2-0442-052c-d9e2c58517fe%28Office.15%29.aspx)|
|[UndoEnabled](http://msdn.microsoft.com/library/c7164cb6-7ce4-b65d-7f5b-1f3987a3fe21%28Office.15%29.aspx)|
|[UserCustomUI](http://msdn.microsoft.com/library/cdd28d78-a75a-b8c4-71e9-74c24ee9ecf1%28Office.15%29.aspx)|
|[Validation](http://msdn.microsoft.com/library/725533ed-49bd-5796-972c-9e84896a3139%28Office.15%29.aspx)|
|[VBProject](http://msdn.microsoft.com/library/087e9cdc-c21d-6f02-05ce-4c3fa6e09cff%28Office.15%29.aspx)|
|[VBProjectData](http://msdn.microsoft.com/library/dca456ea-dc82-0092-35d1-68b95d51e0b2%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/336b6825-3d1c-9589-e916-f8d7821f6383%28Office.15%29.aspx)|
|[ZoomBehavior](http://msdn.microsoft.com/library/5507fc17-957a-ab7f-d15f-43ad3e8327c6%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/944f11be-053c-7749-178c-5e8b79a32ea8%28Office.15%29.aspx)|

