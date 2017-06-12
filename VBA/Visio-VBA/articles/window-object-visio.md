---
title: Window Object (Visio)
keywords: vis_sdr.chm10305
f1_keywords:
- vis_sdr.chm10305
ms.prod: visio
api_name:
- Visio.Window
ms.assetid: 5b49eb0f-07ea-00c7-52f1-2a3115a4b8ae
ms.date: 06/08/2017
---


# Window Object (Visio)

Represents an open window in a Microsoft Visio instance.


## Remarks

The default property of a  **Window** object is **Application**.

To retrieve


- the active window in an instance of Visio, use the  **ActiveWindow** property of an **Application** object.
    
- a  **Page** object that represents the page shown in the window, use the **Page** property of a **Window** object.
    
- a  **Document** object that represents the document displayed in that window, use the **Document** property.
    
- a  **Selection** object that represents the shapes selected in that window, use the **Selection** property.
    

 **Note**  Beginning with Microsoft Visio 2002, the following methods of the  **Window** object are obsolete: **AddToGroup**, **Cut**, **Combine**, **Copy**, **Delete**, **Duplicate**, **Fragment**, **Group**, **Intersect**, **Join RemoveFromGroup**, **Subtract**, **Trim**, and **Union**. Existing solutions that invoke these methods will continue to work properly; however, new or rebuilt solutions should use these methods with the **Selection** object.

In addition, the  **Window** object's **Paste** method is now obsolete. Use the **Paste** or **PasteSpecial** method of the **Page**, **Master**, or **Shape** object. (Use the **Shape** object in the case of group shapes.)


## Events



|**Name**|
|:-----|
|[BeforeWindowClosed](http://msdn.microsoft.com/library/4543e237-6b2c-d02c-66df-9f90b0266e4b%28Office.15%29.aspx)|
|[BeforeWindowPageTurn](http://msdn.microsoft.com/library/818dd4c6-49bd-37ee-c488-e8e0b33b3968%28Office.15%29.aspx)|
|[BeforeWindowSelDelete](http://msdn.microsoft.com/library/450bd22a-ceef-dcf4-90c0-b7511c3506dc%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/70f7d929-5907-e125-1a7f-b68046c6b9dd%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/8e1aa642-0706-4bdd-1401-d08c190e27e5%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/b0301a71-774b-f256-93eb-d5a3ff523def%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/9bffeab4-9df5-a100-2b30-00ea445e6650%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/97f6aece-2d09-b0cc-3197-c16b7cc976a7%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/cb341aa4-9295-4460-53d7-8770e1534707%28Office.15%29.aspx)|
|[OnKeystrokeMessageForAddon](http://msdn.microsoft.com/library/88f72b93-6ec3-2fd1-cc78-c18f82f1b13d%28Office.15%29.aspx)|
|[QueryCancelWindowClose](http://msdn.microsoft.com/library/42b2533a-7958-affc-c722-8b15a396908f%28Office.15%29.aspx)|
|[SelectionChanged](http://msdn.microsoft.com/library/52f5dc68-51d8-7ee0-a31e-ba7525d9c470%28Office.15%29.aspx)|
|[ViewChanged](http://msdn.microsoft.com/library/a65a8e2c-23d5-c582-cd42-4d6f4801d541%28Office.15%29.aspx)|
|[WindowActivated](http://msdn.microsoft.com/library/8fc9f6fc-e391-c3f5-ff73-c58acc738bd1%28Office.15%29.aspx)|
|[WindowChanged](http://msdn.microsoft.com/library/ee7e4871-26ca-ea4e-1c7b-2e597d92e143%28Office.15%29.aspx)|
|[WindowCloseCanceled](http://msdn.microsoft.com/library/bef37fff-5c47-9a61-4b84-ee87912d6478%28Office.15%29.aspx)|
|[WindowTurnedToPage](http://msdn.microsoft.com/library/f1f92687-41b3-fc58-d862-93d4343c5808%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/e34a74e0-8a47-a0bb-4ac5-6fdc8c9e5e08%28Office.15%29.aspx)|
|[CenterViewOnShape](http://msdn.microsoft.com/library/23f219be-bfb7-0f5b-89c0-855093e4bbd9%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/43cb221f-ea65-c12a-e664-0f0fb35685e0%28Office.15%29.aspx)|
|[DeselectAll](http://msdn.microsoft.com/library/926c8578-4c78-efbc-d189-b513fee7ee2f%28Office.15%29.aspx)|
|[DockedStencils](http://msdn.microsoft.com/library/d16865ee-a21f-75c7-55c4-8b30f1ae91e3%28Office.15%29.aspx)|
|[GetViewRect](http://msdn.microsoft.com/library/3281d1af-6745-1b74-5071-e388fa1dc32c%28Office.15%29.aspx)|
|[GetWindowRect](http://msdn.microsoft.com/library/272714c6-3502-4baa-5006-2dcec8c0dfbd%28Office.15%29.aspx)|
|[NewWindow](http://msdn.microsoft.com/library/0cca00d4-9cf4-6a30-b9f2-a37fbad69296%28Office.15%29.aspx)|
|[Scroll](http://msdn.microsoft.com/library/7d30ce8f-03b1-26ff-1495-efb9213083fa%28Office.15%29.aspx)|
|[ScrollViewTo](http://msdn.microsoft.com/library/c2930ee2-f56f-2e3e-cc9a-db73e1d99cd1%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/04394905-0b6b-a07d-4085-a46cecf8afe3%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/81c32217-3336-3017-ecdc-cfa0f6048fc2%28Office.15%29.aspx)|
|[SetViewRect](http://msdn.microsoft.com/library/ab2da074-6e55-3aa7-9c4a-ae299b65a9c9%28Office.15%29.aspx)|
|[SetWindowRect](http://msdn.microsoft.com/library/f9f24c79-9c8f-ec0d-f894-1c10150db75e%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AllowEditing](http://msdn.microsoft.com/library/805ed8a9-1835-0d7b-9bbe-717ff21af3c9%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/2cde63bb-7e4b-c4e7-5be4-ba55d31c5545%28Office.15%29.aspx)|
|[BackgroundColor](http://msdn.microsoft.com/library/5c954890-aa8f-7dc7-c64c-38fd8f3317fe%28Office.15%29.aspx)|
|[BackgroundColorGradient](http://msdn.microsoft.com/library/a23e1075-9a3f-e04a-c6eb-8e4d983b8970%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/be7ee0b3-2891-d5e1-b196-13071ccb2edb%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/305471a6-6497-34b4-dfd5-ff37ccb59fff%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/18421210-d799-dc45-e7e3-39fe5c7f4c09%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/bf05dfe0-b6c0-1ea9-7ce4-af2bd98bbecd%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/b430959b-b7b1-e4a1-d638-6f3ce30e5129%28Office.15%29.aspx)|
|[InPlace](http://msdn.microsoft.com/library/2784b807-0d66-e1db-4936-1b552c06d46b%28Office.15%29.aspx)|
|[IsEditingOLE](http://msdn.microsoft.com/library/aa65ed76-b381-e642-7a29-327b50bc5737%28Office.15%29.aspx)|
|[IsEditingText](http://msdn.microsoft.com/library/2db084a6-8d07-3d29-f3c3-6f19fe50dfab%28Office.15%29.aspx)|
|[Master](http://msdn.microsoft.com/library/caf28e17-797a-91b2-c685-27ad0addddfd%28Office.15%29.aspx)|
|[MasterShortcut](http://msdn.microsoft.com/library/ba25a8a7-fdda-4e25-aea6-75332fe90010%28Office.15%29.aspx)|
|[MergeCaption](http://msdn.microsoft.com/library/19461100-0242-28b1-60bc-9b7f2da3af02%28Office.15%29.aspx)|
|[MergeClass](http://msdn.microsoft.com/library/9ab7b4e7-9be3-9cfe-3a45-57825930ca15%28Office.15%29.aspx)|
|[MergeID](http://msdn.microsoft.com/library/473baaa6-ea88-46f3-3d5f-501f280792a3%28Office.15%29.aspx)|
|[MergePosition](http://msdn.microsoft.com/library/0856bcec-191d-5c9c-f44a-cd430bc3ceb8%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/0c557bcd-ee1f-a094-4248-71fed3dffd58%28Office.15%29.aspx)|
|[Page](http://msdn.microsoft.com/library/17474ce8-f2d7-40c7-7882-30257803c81a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/e52a91c1-299d-91c1-1bea-59609d20a24a%28Office.15%29.aspx)|
|[ParentWindow](http://msdn.microsoft.com/library/923c5f95-8cae-3901-ac03-d8e7668a5b7d%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/ba1884f3-27a3-5c0c-5ebb-85d02c773235%28Office.15%29.aspx)|
|[ReviewerMarkupVisible](http://msdn.microsoft.com/library/7b13a89c-4835-93cc-aece-fcbad1a7ed22%28Office.15%29.aspx)|
|[ScrollLock](http://msdn.microsoft.com/library/08459237-ff58-cd39-319f-60d7bb730487%28Office.15%29.aspx)|
|[SelectedCell](http://msdn.microsoft.com/library/104a2b2d-eb12-2917-6332-9a60e4623e74%28Office.15%29.aspx)|
|[SelectedDataRecordset](http://msdn.microsoft.com/library/89c6b4ba-fb39-8932-1fe0-9a8aa2cbaef0%28Office.15%29.aspx)|
|[SelectedDataRowID](http://msdn.microsoft.com/library/8ed4a690-c96f-c134-5b84-459938bd39e8%28Office.15%29.aspx)|
|[SelectedMasters](http://msdn.microsoft.com/library/8a4546b4-4930-8c69-9df6-84e6b5a1bce0%28Office.15%29.aspx)|
|[SelectedText](http://msdn.microsoft.com/library/75397f73-192b-7683-2a46-016d9b458879%28Office.15%29.aspx)|
|[SelectedValidationIssue](http://msdn.microsoft.com/library/7955338a-2a54-2726-a17a-81acc6bcfce7%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/67c3b3d3-9fe4-ff0c-db94-4a2109f29736%28Office.15%29.aspx)|
|[SelectionForDragCopy](http://msdn.microsoft.com/library/e34de916-5dc4-b9af-70b3-7c68340e2afb%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/ee30f9e5-dd79-83c3-5445-eca53b32822f%28Office.15%29.aspx)|
|[ShowConnectPoints](http://msdn.microsoft.com/library/e69f8fc7-243e-0443-4486-7c0db3a532e2%28Office.15%29.aspx)|
|[ShowGrid](http://msdn.microsoft.com/library/288e1b14-5ad5-da14-8f5b-747212093247%28Office.15%29.aspx)|
|[ShowGuides](http://msdn.microsoft.com/library/875bbdb6-c628-d4be-85d8-fc2529b53627%28Office.15%29.aspx)|
|[ShowPageBreaks](http://msdn.microsoft.com/library/8cdfed9b-bca1-062e-ed69-dfb9e8960a9d%28Office.15%29.aspx)|
|[ShowPageOutline](http://msdn.microsoft.com/library/0e1f0413-1619-0e4f-ad44-e810ee2a38d1%28Office.15%29.aspx)|
|[ShowPageTabs](http://msdn.microsoft.com/library/7ce8bf16-6f99-11fe-8c89-637eec507e2f%28Office.15%29.aspx)|
|[ShowRulers](http://msdn.microsoft.com/library/857dc23b-3687-2b52-db6e-358d32a422fa%28Office.15%29.aspx)|
|[ShowScrollBars](http://msdn.microsoft.com/library/46be2c47-d9b0-c3d8-6f8b-cc728feb4ccb%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/4b83c5ab-8c3d-6477-7127-d1a3ec179c2d%28Office.15%29.aspx)|
|[SubType](http://msdn.microsoft.com/library/3e20338f-a63b-462c-731f-4790042b76cb%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/92dd1e1e-2acc-d918-aab6-f267ecc18c26%28Office.15%29.aspx)|
|[ViewFit](http://msdn.microsoft.com/library/5ee12ad7-4acf-aaf9-a928-93fc473e1c8f%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/e713d0cd-def0-0ae2-08c9-fcfed9ffe883%28Office.15%29.aspx)|
|[WindowHandle32](http://msdn.microsoft.com/library/e766aaab-4b6b-2c8b-3ca2-832fae7e38b0%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/6e063a03-71e5-d2e2-d9d0-38fcae604d26%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/71578934-5d04-8e14-6d87-6871a31f9c4e%28Office.15%29.aspx)|
|[Zoom](http://msdn.microsoft.com/library/35b6973f-ede6-e731-acf0-59ef03456c47%28Office.15%29.aspx)|
|[ZoomBehavior](http://msdn.microsoft.com/library/bceab6cf-cad4-58d6-685d-e14950105048%28Office.15%29.aspx)|
|[ZoomLock](http://msdn.microsoft.com/library/9f962982-27e0-a427-de5f-ed4d3ee04e73%28Office.15%29.aspx)|

