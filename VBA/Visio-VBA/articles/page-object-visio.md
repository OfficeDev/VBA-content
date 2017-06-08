---
title: Page Object (Visio)
keywords: vis_sdr.chm10190
f1_keywords:
- vis_sdr.chm10190
ms.prod: visio
api_name:
- Visio.Page
ms.assetid: 7a7f37ab-b448-eb70-b4f1-c185dfbd511e
ms.date: 06/08/2017
---


# Page Object (Visio)

Represents a drawing page, which can be either a foreground page or a background page.


## Remarks

The default property of a  **Page** object is **Name**.

To retrieve the active page in an instance, use the  **ActivePage** property of an **Application** object.

The members of a  **Document** object's **Pages** collection represent the pages in that document. To retrieve a page's shapes, use the **Shapes** property of a **Page** object.


## Events



|**Name**|
|:-----|
|[AfterReplaceShapes](http://msdn.microsoft.com/library/e4005987-acb1-78d7-91fb-c3c2d5b036e3%28Office.15%29.aspx)|
|[BeforePageDelete](http://msdn.microsoft.com/library/4ef3f16a-b393-fa68-1292-7499ffc302c3%28Office.15%29.aspx)|
|[BeforeReplaceShapes](http://msdn.microsoft.com/library/57ea9836-74dd-77c2-6541-f8f61b89c0b6%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/19bec7f7-9813-bbb4-edf1-117b582ce735%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/7753946d-a986-e89e-aac3-d56556b6c84f%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/8d121852-dd5b-45d1-dee6-f838a2533243%28Office.15%29.aspx)|
|[CalloutRelationshipAdded](http://msdn.microsoft.com/library/b5181cd5-e763-a25c-abdc-3b32d2c902a0%28Office.15%29.aspx)|
|[CalloutRelationshipDeleted](http://msdn.microsoft.com/library/06ab7df2-c2a9-2b86-4dd3-817f56dddf6c%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/78c9bc15-6d4b-1580-3d36-2109364a4a1c%28Office.15%29.aspx)|
|[ConnectionsAdded](http://msdn.microsoft.com/library/62495ee5-b2f8-bbe3-cb7f-2b02622a5c13%28Office.15%29.aspx)|
|[ConnectionsDeleted](http://msdn.microsoft.com/library/7be3ec10-0715-8daa-a021-c7e6780c223a%28Office.15%29.aspx)|
|[ContainerRelationshipAdded](http://msdn.microsoft.com/library/4cd95f23-baaa-3987-05f3-a379670efd02%28Office.15%29.aspx)|
|[ContainerRelationshipDeleted](http://msdn.microsoft.com/library/2c56eb44-9a5b-49a7-9137-8bff7d0399af%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/c44afba7-eeb5-3760-7ab3-1e5e86d92060%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/3ab03e1c-e2c1-314b-5f09-853b170096d1%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/ae7bd6b5-8975-26a2-86af-ff12eaef5ebb%28Office.15%29.aspx)|
|[PageChanged](http://msdn.microsoft.com/library/e42dd83e-9d2b-93f7-fe18-e3651fcfa608%28Office.15%29.aspx)|
|[PageDeleteCanceled](http://msdn.microsoft.com/library/5fa17e8b-5c80-962b-482e-f9c46f543a65%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/a9dc79ef-2a4c-398a-4bf3-d29e0cf916f4%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/ee70861c-ca8e-0cc8-ddc4-40c5bcb9f74e%28Office.15%29.aspx)|
|[QueryCancelPageDelete](http://msdn.microsoft.com/library/f862d9ac-c052-31df-9d9a-0ecd8352467a%28Office.15%29.aspx)|
|[QueryCancelReplaceShapes](http://msdn.microsoft.com/library/17ead23f-825a-c608-3315-e2eed6784cd5%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/85ece21a-03b0-d4ff-fb72-b701b0753f1d%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/ab03af54-dd9a-03ca-18ac-e76ca103035b%28Office.15%29.aspx)|
|[ReplaceShapesCanceled](http://msdn.microsoft.com/library/867b1fc1-96bd-cbeb-fd61-b02a96e039ca%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/24e893c8-093e-c846-a74d-12f10c1009e6%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/49ef8516-43bb-b410-5e6c-6903c2bf32fa%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/bc66eadc-21bc-7f17-6878-fddd9aaff855%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/cc831cfe-a0b5-58c8-a204-21a11de4262f%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/ba9a4dcf-db2b-bca4-8c4a-bf7d9234dbb2%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/fd3d6512-2cc6-e7ab-f0dd-c44ee5054890%28Office.15%29.aspx)|
|[ShapeLinkAdded](http://msdn.microsoft.com/library/3d49ffc4-9d08-c228-ba3c-d4d97362bb62%28Office.15%29.aspx)|
|[ShapeLinkDeleted](http://msdn.microsoft.com/library/e19709c4-45e4-f0f1-8e59-72b1ccbdf130%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/656e38cc-3900-86ba-1f1e-bfcc5b3697c7%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/c3b5ea4c-0552-5bea-1bf5-6abd47d1fc63%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/39e22317-9189-29b0-035a-404cd67844c6%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddGuide](http://msdn.microsoft.com/library/7be0cc07-6322-a3f0-3292-6dc66804db44%28Office.15%29.aspx)|
|[AutoConnectMany](http://msdn.microsoft.com/library/292d0f58-d753-6ef3-fd62-269fd44d003c%28Office.15%29.aspx)|
|[AutoSizeDrawing](http://msdn.microsoft.com/library/00ae0d14-3268-f6d5-2adb-4653958b6eee%28Office.15%29.aspx)|
|[AvoidPageBreaks](http://msdn.microsoft.com/library/70e99d9d-cce0-c162-5836-0a68e375e4c3%28Office.15%29.aspx)|
|[BoundingBox](http://msdn.microsoft.com/library/f281e304-057f-5555-8efd-fd81d088b8cd%28Office.15%29.aspx)|
|[CenterDrawing](http://msdn.microsoft.com/library/9e5f7c27-f2ef-f8e1-b530-9d8d41960193%28Office.15%29.aspx)|
|[CreateSelection](http://msdn.microsoft.com/library/7bd29416-d6b4-d7f9-dd96-2ec66c2d4e6b%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/7adc0e81-7000-2bfa-cca5-c74c3fcbac5c%28Office.15%29.aspx)|
|[DrawArcByThreePoints](http://msdn.microsoft.com/library/dfa20dfd-22f7-6d99-2649-d8401bf93a19%28Office.15%29.aspx)|
|[DrawBezier](http://msdn.microsoft.com/library/49cf1bfb-5b88-ca8d-4451-a9884768f780%28Office.15%29.aspx)|
|[DrawCircularArc](http://msdn.microsoft.com/library/2c57ec5d-418c-df3b-a599-61d5fa560467%28Office.15%29.aspx)|
|[DrawLine](http://msdn.microsoft.com/library/a03308a6-7ad0-ecaa-d15d-a243402c8bd3%28Office.15%29.aspx)|
|[DrawNURBS](http://msdn.microsoft.com/library/f3c7e6fe-71a4-4809-b60a-a34cebd737b1%28Office.15%29.aspx)|
|[DrawOval](http://msdn.microsoft.com/library/9e3afc60-b14d-c831-5271-be782366a2d6%28Office.15%29.aspx)|
|[DrawPolyline](http://msdn.microsoft.com/library/406ac09e-c25f-5de6-1c0b-e2a456ed5ec0%28Office.15%29.aspx)|
|[DrawQuarterArc](http://msdn.microsoft.com/library/f1d658cf-62de-5979-bd0c-0eea54fb08c4%28Office.15%29.aspx)|
|[DrawRectangle](http://msdn.microsoft.com/library/3ace50fe-cc78-1412-28d6-5bc1dbe73700%28Office.15%29.aspx)|
|[DrawSpline](http://msdn.microsoft.com/library/a75d7f02-5bfd-f341-ca24-06762e56aca3%28Office.15%29.aspx)|
|[Drop](http://msdn.microsoft.com/library/015615a8-fe64-5b76-39ba-ef7ed62e6846%28Office.15%29.aspx)|
|[DropCallout](http://msdn.microsoft.com/library/72edbd4b-e068-6dac-0298-9f746a728892%28Office.15%29.aspx)|
|[DropConnected](http://msdn.microsoft.com/library/7e16dc46-df74-4482-91a4-b0a115f979b2%28Office.15%29.aspx)|
|[DropContainer](http://msdn.microsoft.com/library/14da134d-6a3f-25c3-37c4-eb8b51c213ab%28Office.15%29.aspx)|
|[DropIntoList](http://msdn.microsoft.com/library/fcefca11-d64b-9f95-a00e-bf9968d26267%28Office.15%29.aspx)|
|[DropLegend](http://msdn.microsoft.com/library/8253eafd-4d87-9f1c-833c-cb553c1b73cf%28Office.15%29.aspx)|
|[DropLinked](http://msdn.microsoft.com/library/e975a150-ff48-7cae-3e3b-f21f88f2fbd2%28Office.15%29.aspx)|
|[DropMany](http://msdn.microsoft.com/library/81fc5b8d-3152-de69-2f8e-90d530aa5e08%28Office.15%29.aspx)|
|[DropManyLinkedU](http://msdn.microsoft.com/library/0b80591a-a563-bdad-b048-e15693410547%28Office.15%29.aspx)|
|[DropManyU](http://msdn.microsoft.com/library/e61d9e8f-3838-240e-b8da-c5f1d8b3eb12%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/394be23b-997d-0da1-b3bd-8278564fb4e0%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/7eef4f56-4b47-bebc-4657-fcd1a5d5b0db%28Office.15%29.aspx)|
|[GetCallouts](http://msdn.microsoft.com/library/a0300c64-4bdd-e442-c00c-a727debbf6b8%28Office.15%29.aspx)|
|[GetContainers](http://msdn.microsoft.com/library/17d9365b-f9ac-85ba-e1cb-cd02ea1a2f22%28Office.15%29.aspx)|
|[GetFormulas](http://msdn.microsoft.com/library/d501f50f-2e8b-36bb-e303-97f445908e4a%28Office.15%29.aspx)|
|[GetFormulasU](http://msdn.microsoft.com/library/8d7ba7d3-51e6-cd65-78ad-27640188e348%28Office.15%29.aspx)|
|[GetResults](http://msdn.microsoft.com/library/5af0a38f-fdc9-e826-99b0-6090bb372bc1%28Office.15%29.aspx)|
|[GetShapesLinkedToData](http://msdn.microsoft.com/library/3196f7f9-1b7c-8070-444d-c1a55f0c205f%28Office.15%29.aspx)|
|[GetShapesLinkedToDataRow](http://msdn.microsoft.com/library/d305eccc-4121-be3a-a389-f50234e526f1%28Office.15%29.aspx)|
|[GetTheme](http://msdn.microsoft.com/library/31c84e69-0bc8-2d1a-84d8-7397110d74ae%28Office.15%29.aspx)|
|[GetThemeVariant](http://msdn.microsoft.com/library/40c2be31-fdb0-68ee-a129-2788b1b17c82%28Office.15%29.aspx)|
|[Import](http://msdn.microsoft.com/library/a84086c3-694d-8cf3-e6f7-ba84e182dd4a%28Office.15%29.aspx)|
|[InsertFromFile](http://msdn.microsoft.com/library/03762511-9f2f-6691-ac82-dcff74fcde1d%28Office.15%29.aspx)|
|[InsertObject](http://msdn.microsoft.com/library/74081ecf-59ee-44e8-6fc8-3ccc0915e110%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/3611d496-ecb9-674e-b435-8462d55f7256%28Office.15%29.aspx)|
|[LayoutChangeDirection](http://msdn.microsoft.com/library/f818785b-d845-34de-50d1-e68c3c09dda9%28Office.15%29.aspx)|
|[LayoutIncremental](http://msdn.microsoft.com/library/db112261-120d-e2e8-18f0-91b1bba0a3a4%28Office.15%29.aspx)|
|[LinkShapesToDataRows](http://msdn.microsoft.com/library/306c8edf-04ea-1e54-b3cf-63ea0352c242%28Office.15%29.aspx)|
|[OpenDrawWindow](http://msdn.microsoft.com/library/b5c4e800-fdba-2529-1c04-afa261377469%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/73dd3b44-1288-26d1-4956-93f187d71886%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/949a507a-1cc2-0b52-b0dd-3ad40ac9ecdf%28Office.15%29.aspx)|
|[PasteToLocation](http://msdn.microsoft.com/library/d24cc1b3-c0c7-d529-b94f-0fea82d124ef%28Office.15%29.aspx)|
|[Print](http://msdn.microsoft.com/library/021cdd78-1699-4345-5b32-c2c0a300ca00%28Office.15%29.aspx)|
|[PrintTile](http://msdn.microsoft.com/library/221efce0-c706-8583-50a5-ba28ef620fdf%28Office.15%29.aspx)|
|[ResizeToFitContents](http://msdn.microsoft.com/library/26b96288-7d8b-a999-ef45-a586110cc8b9%28Office.15%29.aspx)|
|[SetFormulas](http://msdn.microsoft.com/library/141de5db-67dc-11c9-69a1-29601bf71cb1%28Office.15%29.aspx)|
|[SetResults](http://msdn.microsoft.com/library/2f50a50c-3223-4948-e802-af97d1b2e815%28Office.15%29.aspx)|
|[SetTheme](http://msdn.microsoft.com/library/5a186f58-9a7a-bd8a-826b-85da75a4d59f%28Office.15%29.aspx)|
|[SetThemeVariant](http://msdn.microsoft.com/library/8393a95f-83ca-0efa-d987-ae498bfe5e9d%28Office.15%29.aspx)|
|[ShapeIDsToUniqueIDs](http://msdn.microsoft.com/library/b89e82db-3c7b-fb73-2f4c-10056c6e7b28%28Office.15%29.aspx)|
|[SplitConnector](http://msdn.microsoft.com/library/b2d371b5-3769-00cd-688f-2391a8c504e9%28Office.15%29.aspx)|
|[UniqueIDsToShapeIDs](http://msdn.microsoft.com/library/86d0d47c-d356-04ba-51ce-7d682fd165ae%28Office.15%29.aspx)|
|[VisualBoundingBox](http://msdn.microsoft.com/library/95e8a977-55c9-307a-bade-120cb8acdf9b%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/e4f0a4ad-d99c-efec-d4e9-8a5fc625288e%28Office.15%29.aspx)|
|[AutoSize](http://msdn.microsoft.com/library/777155fb-21a6-f7d2-3eef-66ed09a00628%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/fee785fd-2872-a64e-a80e-46034255b414%28Office.15%29.aspx)|
|[BackPage](http://msdn.microsoft.com/library/cef2dac4-cf12-d692-cbbc-a6023f2d78e0%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/9618c86c-96c0-be95-ee20-5d1b99f4d5e8%28Office.15%29.aspx)|
|[Connects](http://msdn.microsoft.com/library/55b98c54-0507-c87b-a983-b06e0fcc707d%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/3616486c-4c54-698f-19ff-ddde2f5e7bec%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/7841962e-c2c5-0cf3-2073-fc97a050e32e%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/61904830-7949-98c0-eb69-a6d685b3a38c%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/00bc8738-ad54-a5ae-a6aa-bfb762ee0fa7%28Office.15%29.aspx)|
|[Layers](http://msdn.microsoft.com/library/62e3aae6-1cb1-695e-81ec-eabdd6b44ef9%28Office.15%29.aspx)|
|[LayoutRoutePassive](http://msdn.microsoft.com/library/7244abb5-0c8f-d68b-4b2d-3e192afe1d80%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/745bb4cf-b79c-4212-325b-40b4e1c9bc81%28Office.15%29.aspx)|
|[NameU](http://msdn.microsoft.com/library/d4e8c719-8667-caaa-3a41-1f80ec65fd75%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/54da9c26-fffe-7121-81e7-3a883d103edd%28Office.15%29.aspx)|
|[OLEObjects](http://msdn.microsoft.com/library/8546ecb2-4889-465f-af6c-c312b1b4900a%28Office.15%29.aspx)|
|[OriginalPage](http://msdn.microsoft.com/library/4c4ca104-755a-8092-51e9-b78a6e45c95b%28Office.15%29.aspx)|
|[PageSheet](http://msdn.microsoft.com/library/495709a8-92f0-6fdf-753f-7ac25c5daaab%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/5e4fb8d6-bb4e-dce9-a516-3bf0f0746e82%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/2e70f00f-6f42-4449-2fcf-ec79f0097296%28Office.15%29.aspx)|
|[PrintTileCount](http://msdn.microsoft.com/library/f15eff27-1d20-7151-e773-1ab4de4161db%28Office.15%29.aspx)|
|[ReviewerID](http://msdn.microsoft.com/library/f3de7746-f1f7-4a94-6fcb-e3c2775ed748%28Office.15%29.aspx)|
|[ShapeComments](http://msdn.microsoft.com/library/b7d86594-ba1f-627b-222f-905da1b1201e%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/b6a5c174-c1d6-049b-8aec-8337c47341d7%28Office.15%29.aspx)|
|[SpatialSearch](http://msdn.microsoft.com/library/539d2884-2092-6eb5-8d22-af8062f139db%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/791e19c4-7524-2370-652d-f4377e09357f%28Office.15%29.aspx)|
|[ThemeColors](http://msdn.microsoft.com/library/a3f4bc4e-3dbb-9d50-9d71-f77b39ec0ac3%28Office.15%29.aspx)|
|[ThemeEffects](http://msdn.microsoft.com/library/566ee9aa-9c45-e53b-2634-c666565e6fbb%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/7e9c949d-11a6-b9c4-6d25-bc70e8ec9034%28Office.15%29.aspx)|

