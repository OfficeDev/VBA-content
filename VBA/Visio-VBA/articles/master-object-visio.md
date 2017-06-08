---
title: Master Object (Visio)
keywords: vis_sdr.chm10130
f1_keywords:
- vis_sdr.chm10130
ms.prod: visio
api_name:
- Visio.Master
ms.assetid: 1a69e4d7-2b72-f712-d36c-c565af64c278
ms.date: 06/08/2017
---


# Master Object (Visio)

Represents a master in a stencil.


## Remarks

You retrieve a particular  **Master** object from the **Masters** collection of a **Document** object whose stencil contains that master.

The default property of a  **Master** object is **Name**.

To create an instance of a master in a drawing, use the  **Drop** method of a **Page** object that represents a drawing page.


## Events



|**Name**|
|:-----|
|[BeforeMasterDelete](http://msdn.microsoft.com/library/46b455db-9165-0ed4-ebf3-15e1794313be%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/e2f86944-6ca2-6535-ee08-889af9694fd6%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/21921e16-3e05-6232-ed89-76217b76149f%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/1d39001d-6efa-7d58-1eaa-f6c2531e2018%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/53323234-8e92-de8b-65b8-20eb867748dd%28Office.15%29.aspx)|
|[ConnectionsAdded](http://msdn.microsoft.com/library/15c772fe-d5fb-901e-f1d4-1d3eb0cb7c64%28Office.15%29.aspx)|
|[ConnectionsDeleted](http://msdn.microsoft.com/library/dc043012-d653-8f37-372e-f7532047aa81%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/b585e434-fd81-93ae-92a6-5cc1d21c1afa%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/6d2a9ab6-778e-cbba-0b63-f7d38116dc85%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/ec87e679-2b8f-de85-81b9-ccb4a9df7ae2%28Office.15%29.aspx)|
|[MasterChanged](http://msdn.microsoft.com/library/922120cc-56e0-143b-7a8b-754bc368af47%28Office.15%29.aspx)|
|[MasterDeleteCanceled](http://msdn.microsoft.com/library/a682fab6-1fc9-65ba-83a1-408d048ee81e%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/c23d7ed0-0ad4-fa20-4b4f-fa453716fbd5%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/37625c3b-49e2-d3ba-5270-2dcb65062f08%28Office.15%29.aspx)|
|[QueryCancelMasterDelete](http://msdn.microsoft.com/library/33690e0f-821e-42cd-ec52-3ade1a1ceadc%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/c85569ca-b802-7a7e-6b24-d89852d2d0bc%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/44ce0f2e-e877-ec7f-b5ec-1c3ff3b9749a%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/c004e65c-1770-edf1-9d1e-a1a02a15fc39%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/87ecdfcb-616f-0b47-bfa4-216ef456deaa%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/d679f866-c939-faff-d8da-cdddb2131054%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/e1a2a7bf-bfe1-acfc-ae04-308f9fda7c0a%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/74eb2604-bcb2-0cba-37e2-50ad896991ca%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/401f6d32-d1fb-f019-52a3-d553b8516ecf%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/37de7351-969b-5b24-fde2-e4473e92b344%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/9224577c-a285-c26f-60be-3adbf3285ef3%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/0bbe537e-9bae-62a9-7e29-aea71ab3c8f9%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddGuide](http://msdn.microsoft.com/library/7beba614-244b-f559-50c7-5156ca4510b1%28Office.15%29.aspx)|
|[BoundingBox](http://msdn.microsoft.com/library/23ef5e08-fcb4-93e6-2ed5-818d34f99a8e%28Office.15%29.aspx)|
|[CenterDrawing](http://msdn.microsoft.com/library/1bf660a3-30eb-4a0b-fcea-66d0e0574ae0%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/69607a2c-dc59-d170-733a-3557a996a67e%28Office.15%29.aspx)|
|[CreateSelection](http://msdn.microsoft.com/library/52db8b1b-e253-549f-c3ba-d661fa7b675e%28Office.15%29.aspx)|
|[CreateShortcut](http://msdn.microsoft.com/library/e808ba09-b85a-52bb-55e2-ced37f426a3b%28Office.15%29.aspx)|
|[DataGraphicDelete](http://msdn.microsoft.com/library/aa84af70-975c-3747-1976-b872a6c2fa36%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/8f71e69e-7d7d-7732-738c-ad262b0367ae%28Office.15%29.aspx)|
|[DrawArcByThreePoints](http://msdn.microsoft.com/library/d2df1c41-8164-d941-21a8-2e1b00de6199%28Office.15%29.aspx)|
|[DrawBezier](http://msdn.microsoft.com/library/4cbefabf-530e-2c6d-0751-45efa2bb0980%28Office.15%29.aspx)|
|[DrawCircularArc](http://msdn.microsoft.com/library/f9557127-8470-2968-3056-0e295cd05633%28Office.15%29.aspx)|
|[DrawLine](http://msdn.microsoft.com/library/c29810a2-c1eb-82cc-ab19-236a89baf7b0%28Office.15%29.aspx)|
|[DrawNURBS](http://msdn.microsoft.com/library/7dcfef4a-5b69-9a8b-3966-9b3089bdaac3%28Office.15%29.aspx)|
|[DrawOval](http://msdn.microsoft.com/library/092a59d6-1b43-c094-e2ae-480ee7b32b73%28Office.15%29.aspx)|
|[DrawPolyline](http://msdn.microsoft.com/library/a599e60c-ccd6-ce6b-7e54-f65f8500447d%28Office.15%29.aspx)|
|[DrawQuarterArc](http://msdn.microsoft.com/library/6c728c0c-8317-6114-70e5-e5cb68a5729f%28Office.15%29.aspx)|
|[DrawRectangle](http://msdn.microsoft.com/library/e41ec411-ccd7-0fe6-f560-cf3934d18b59%28Office.15%29.aspx)|
|[DrawSpline](http://msdn.microsoft.com/library/a255978d-5479-ba7e-4520-0a8d18390ea6%28Office.15%29.aspx)|
|[Drop](http://msdn.microsoft.com/library/13abc8fc-7b3c-98cf-3965-3ac7b3d15e85%28Office.15%29.aspx)|
|[DropMany](http://msdn.microsoft.com/library/fb0ef035-c1ce-5703-e2e8-0f9b63b186bf%28Office.15%29.aspx)|
|[DropManyU](http://msdn.microsoft.com/library/467356ff-d2d9-71d9-d533-b88099bf2fae%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/212bcc8e-646c-37df-9387-4605b72b6edd%28Office.15%29.aspx)|
|[ExportIcon](http://msdn.microsoft.com/library/8b13f92f-537a-1efb-b2b0-531a8054e89b%28Office.15%29.aspx)|
|[GetFormulas](http://msdn.microsoft.com/library/09ee33a3-41fc-3ac2-4f5e-1e857f685049%28Office.15%29.aspx)|
|[GetFormulasU](http://msdn.microsoft.com/library/d5a419e2-9630-a724-af44-f2f1b0166c80%28Office.15%29.aspx)|
|[GetResults](http://msdn.microsoft.com/library/d532a2ed-2246-8c90-2d77-df2df05a395f%28Office.15%29.aspx)|
|[Import](http://msdn.microsoft.com/library/3b13025f-1a83-0dcf-41e1-03cd83dfc7be%28Office.15%29.aspx)|
|[ImportIcon](http://msdn.microsoft.com/library/886d724d-9d02-ab6f-8049-80fa04f8caec%28Office.15%29.aspx)|
|[InsertFromFile](http://msdn.microsoft.com/library/5a24e289-675a-d08b-36f7-0cfaedac5aaf%28Office.15%29.aspx)|
|[InsertObject](http://msdn.microsoft.com/library/7b663eef-ed40-486b-2b5b-e7c7066c2300%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/acab2dc3-daf8-57c2-cbf8-edf647a12a09%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/3f14f3b2-1cfb-ccf9-b344-7fbf80ae9a26%28Office.15%29.aspx)|
|[OpenDrawWindow](http://msdn.microsoft.com/library/5f17d4a0-6b5d-bb85-cff7-047bd18ff1ee%28Office.15%29.aspx)|
|[OpenIconWindow](http://msdn.microsoft.com/library/5e2b2437-05cc-4855-e0bb-96b097c98d3c%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/ee8a4c79-9a10-d852-70d3-4856627efb8a%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/6ca1994b-feb4-6b0d-c2c4-8a134eb284f1%28Office.15%29.aspx)|
|[PasteToLocation](http://msdn.microsoft.com/library/c5c94265-23ee-5516-525d-ed3f34d2e7bf%28Office.15%29.aspx)|
|[ResizeToFitContents](http://msdn.microsoft.com/library/982fa4c4-014c-319d-a73e-f6bbc28f16e8%28Office.15%29.aspx)|
|[SetFormulas](http://msdn.microsoft.com/library/fb419eb5-6bd3-cfc7-d358-cef9e68dddbf%28Office.15%29.aspx)|
|[SetResults](http://msdn.microsoft.com/library/6be7dd71-55a7-777c-e1b7-8f41c028e843%28Office.15%29.aspx)|
|[VisualBoundingBox](http://msdn.microsoft.com/library/478d636f-e741-cf6b-3e16-b5faf70a9f14%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AlignName](http://msdn.microsoft.com/library/5df055eb-ddb1-2d2a-1d94-93781960b3a9%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/88b2fd6e-8f7e-3caa-5316-35a6a0060793%28Office.15%29.aspx)|
|[BaseID](http://msdn.microsoft.com/library/85ca3c0d-5015-b303-7102-144768acb6a8%28Office.15%29.aspx)|
|[Connects](http://msdn.microsoft.com/library/72c01ae0-9134-d384-b860-dbb333a498fe%28Office.15%29.aspx)|
|[DataGraphicHidden](http://msdn.microsoft.com/library/adcf1867-8541-785b-d8ad-dd44583473b9%28Office.15%29.aspx)|
|[DataGraphicHidesText](http://msdn.microsoft.com/library/c1a08780-0873-3d8b-1872-edc8a6515840%28Office.15%29.aspx)|
|[DataGraphicHorizontalPosition](http://msdn.microsoft.com/library/d9c98a41-ffc0-152e-2150-0915bd38bcac%28Office.15%29.aspx)|
|[DataGraphicShowBorder](http://msdn.microsoft.com/library/203d631c-d838-ea0a-f67a-39de513e738e%28Office.15%29.aspx)|
|[DataGraphicVerticalPosition](http://msdn.microsoft.com/library/779f360e-7529-7fe6-87e7-f41cc9334c83%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/b95000f8-67df-99f4-bbfc-020b14ae73b8%28Office.15%29.aspx)|
|[EditCopy](http://msdn.microsoft.com/library/69d13b8f-c5af-d9c9-b92e-00e6eadf660a%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/02a4d80f-fbc6-6491-5f8b-ce98dd5c2aa8%28Office.15%29.aspx)|
|[GraphicItems](http://msdn.microsoft.com/library/615b4909-c248-3ebd-c7c1-53151464cee9%28Office.15%29.aspx)|
|[Hidden](http://msdn.microsoft.com/library/d28eb888-75d7-bbd2-e6d3-3e412cca85d4%28Office.15%29.aspx)|
|[Icon](http://msdn.microsoft.com/library/2e9c7bbd-d8fd-e932-4a6b-bbd845aef4f0%28Office.15%29.aspx)|
|[IconSize](http://msdn.microsoft.com/library/c6516b30-642d-1e61-22b4-f95d6c47a8ec%28Office.15%29.aspx)|
|[IconUpdate](http://msdn.microsoft.com/library/3978c650-47d5-e961-53c2-d99dd4c2ca7c%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/9064e708-f939-9522-b8f7-24488d780bc0%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/48a90dee-ce11-ef81-e58a-e4a3cdb899dc%28Office.15%29.aspx)|
|[IndexInStencil](http://msdn.microsoft.com/library/3c2c12c4-0233-4aa3-c3d7-a3613bb391ad%28Office.15%29.aspx)|
|[IsChanged](http://msdn.microsoft.com/library/8e557655-3e16-3e96-99a2-b097fa6abd75%28Office.15%29.aspx)|
|[Layers](http://msdn.microsoft.com/library/6c78d629-506c-54aa-e0cc-7fd807cdfffb%28Office.15%29.aspx)|
|[MatchByName](http://msdn.microsoft.com/library/4edb0e5f-7e87-c66d-b842-318cd0eba5d5%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/66ca8cd6-c784-efbb-a2b6-2b0fcce7d5b1%28Office.15%29.aspx)|
|[NameU](http://msdn.microsoft.com/library/87530cb6-5ac1-55c4-9210-9989c5f589c3%28Office.15%29.aspx)|
|[NewBaseID](http://msdn.microsoft.com/library/bee59c61-06de-ebb9-a8aa-599fc788e4e1%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/958b08f3-a52b-d6cb-2360-ca2ddf758e3c%28Office.15%29.aspx)|
|[OLEObjects](http://msdn.microsoft.com/library/b51fbdc2-a236-4733-5a2e-b8e75d457d64%28Office.15%29.aspx)|
|[OneD](http://msdn.microsoft.com/library/917f8cfc-a2fc-7572-936a-69956d139131%28Office.15%29.aspx)|
|[Original](http://msdn.microsoft.com/library/33636aa0-2b2b-9edb-3738-ac193eaab212%28Office.15%29.aspx)|
|[PageSheet](http://msdn.microsoft.com/library/8ec4d38a-79fe-018d-9bc8-3a9c0221f018%28Office.15%29.aspx)|
|[PatternFlags](http://msdn.microsoft.com/library/cf7d5e0e-802e-c65b-6260-eaf68dfe6eb4%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/6840a242-85d8-b93e-242b-90c584a9b422%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/b882b05f-5e54-aab8-db88-1e66cf825581%28Office.15%29.aspx)|
|[Prompt](http://msdn.microsoft.com/library/7467c2dd-5cf6-0af0-bc4d-522889d69707%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/56db5c02-9b55-dfe1-993b-c23e93e84577%28Office.15%29.aspx)|
|[SpatialSearch](http://msdn.microsoft.com/library/d71b05b7-32e1-d3c8-668e-6e96595acd59%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/1cc33fe9-e317-ab3d-1ce1-a7f8c619c4f2%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/4688ff5d-2f9a-fcaf-6a73-0aa50562b24a%28Office.15%29.aspx)|
|[UniqueID](http://msdn.microsoft.com/library/99d0655c-da5c-9d0a-4936-2fa24821e097%28Office.15%29.aspx)|

