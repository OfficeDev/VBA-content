---
title: Pages Object (Visio)
keywords: vis_sdr.chm10195
f1_keywords:
- vis_sdr.chm10195
ms.prod: visio
api_name:
- Visio.Pages
ms.assetid: 45eec568-b5cc-5e80-ff5c-4dfa567efb5d
ms.date: 06/08/2017
---


# Pages Object (Visio)

Includes a  **Page** object for each drawing page in a document.


## Remarks

To retrieve a  **Pages** collection, use the **Pages** property of a **Document** object.

The default property of a  **Pages** collection is **Item**.

The order of items in a  **Pages** collection is significant: if there are _n_ foreground pages in a document, the first _n_ pages in its **Pages** collection are foreground pages and are in order. The remaining pages in the collection are the background pages of the document; these are in no particular order.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVPages**
    

## Events



|**Name**|
|:-----|
|[AfterReplaceShapes](http://msdn.microsoft.com/library/05c33bdd-e697-d36e-46a8-45705e9ad2c2%28Office.15%29.aspx)|
|[BeforePageDelete](http://msdn.microsoft.com/library/52fbea6b-0258-8610-74e2-74ade9f8ae49%28Office.15%29.aspx)|
|[BeforeReplaceShapes](http://msdn.microsoft.com/library/3f6dbc31-0583-dd67-0432-335d6df7a50c%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/2c0ad4cf-f734-f5f2-1fea-c5ce846cfd05%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/e83bb4cc-b9a0-1435-507f-149f5a108ab5%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/3006644c-9c2e-6a35-f484-f2dc3d12c1e3%28Office.15%29.aspx)|
|[CalloutRelationshipAdded](http://msdn.microsoft.com/library/45f350ca-05ed-b775-d5da-b0d9c8a5c885%28Office.15%29.aspx)|
|[CalloutRelationshipDeleted](http://msdn.microsoft.com/library/5e5a3149-9179-8e7c-3728-36e7e2cc3c71%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/eb25f423-76eb-b82a-953b-460ab2b10a00%28Office.15%29.aspx)|
|[ConnectionsAdded](http://msdn.microsoft.com/library/7b2a471c-425f-0ab5-2cae-561dc67e343c%28Office.15%29.aspx)|
|[ConnectionsDeleted](http://msdn.microsoft.com/library/af35574e-2855-2581-e51a-b777eaa83aca%28Office.15%29.aspx)|
|[ContainerRelationshipAdded](http://msdn.microsoft.com/library/8d7480e7-0131-8c02-11ad-d5784679e387%28Office.15%29.aspx)|
|[ContainerRelationshipDeleted](http://msdn.microsoft.com/library/ed72e2e1-00c8-9ae0-eb53-57fe75035345%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/a665309f-bf0c-58b1-35ec-3843ef2a1e77%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/97c8766e-b682-7df9-7e2c-9a558d5d09f1%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/331fc5c3-bd2c-129c-fed2-3f0fe53f95e5%28Office.15%29.aspx)|
|[PageAdded](http://msdn.microsoft.com/library/59268803-17a2-e1fd-70da-45506b9076a3%28Office.15%29.aspx)|
|[PageChanged](http://msdn.microsoft.com/library/7e6f4eec-4043-fa9b-4225-6f5120676bde%28Office.15%29.aspx)|
|[PageDeleteCanceled](http://msdn.microsoft.com/library/72e07c4f-70c9-a310-4086-ba2aff1cafbc%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/97d86952-e77f-55ad-84aa-237ee750f6c9%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/02e97010-02b9-1062-22fb-0b3d24eb90f1%28Office.15%29.aspx)|
|[QueryCancelPageDelete](http://msdn.microsoft.com/library/ca487884-ca7f-a1b6-1800-95550a056c8f%28Office.15%29.aspx)|
|[QueryCancelReplaceShapes](http://msdn.microsoft.com/library/d11ff976-0016-da6b-92fb-379baa7e8f94%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/d9749c36-d336-f251-7f69-c48bcf590d56%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/b1844dea-5b97-2a8e-5ec7-143afdf44067%28Office.15%29.aspx)|
|[ReplaceShapesCanceled](http://msdn.microsoft.com/library/f0ce8c66-7a15-5f91-8c89-e177bc6671d2%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/76ffc5b0-fccb-d963-76cd-fe2fcc9829f2%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/3644b404-e5e5-b18c-5131-406822fd66e1%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/7a68596c-8d8e-255d-0b3a-4490cb2f99d5%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/a012a091-b7cc-0d7c-36a2-bbfc675356d0%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/c96ef86a-2635-2e2b-4d3c-4cb24ceaae69%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/c4af9e02-79ad-0840-2e74-10fa946abd10%28Office.15%29.aspx)|
|[ShapeLinkAdded](http://msdn.microsoft.com/library/432a8daa-9545-0df7-3e78-65464e74c7df%28Office.15%29.aspx)|
|[ShapeLinkDeleted](http://msdn.microsoft.com/library/f39e1e75-3f1a-04a7-6232-8d1d17560175%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/9a566e9f-479f-c69d-8831-21fd7694c201%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/612fac07-8abe-4697-9634-108eeea78f0e%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/9ee7c970-7cb4-3683-b71c-1c828bbd4ec4%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/b2e09b89-4232-fffe-28b2-ceb468dd2837%28Office.15%29.aspx)|
|[GetNames](http://msdn.microsoft.com/library/9e3c9e6a-94fe-aa1f-0591-60e6f7314b7f%28Office.15%29.aspx)|
|[GetNamesU](http://msdn.microsoft.com/library/eb7ac155-5124-f25d-3c5a-a30773940dd0%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/f3f8fdf7-8ca2-aa43-a0eb-3fd5151ad8da%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/1e240cc4-07f3-ceb1-7eb3-7a6d5071f630%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/d2825f21-f4ba-05d6-62b8-646e8c4be43e%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/2baa8080-d099-c2c0-86f6-040c8edd82c0%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/c52ace02-486f-d50b-caf5-109b78008d77%28Office.15%29.aspx)|
|[ItemFromID](http://msdn.microsoft.com/library/0355a186-b7bf-51e5-bb2c-433417cf2d33%28Office.15%29.aspx)|
|[ItemU](http://msdn.microsoft.com/library/cb5af44e-b8de-229d-b7da-d6377f68c494%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/b36f235d-2c04-8d11-e50a-59c245c2fc0b%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/fb239aaf-ff62-8231-dd47-4fe8b70b3062%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/429cc898-4daf-e269-4e10-ac808f429d62%28Office.15%29.aspx)|

