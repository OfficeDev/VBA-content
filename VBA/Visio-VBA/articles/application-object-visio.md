---
title: Application Object (Visio)
keywords: vis_sdr.chm10040
f1_keywords:
- vis_sdr.chm10040
ms.prod: visio
api_name:
- Visio.Application
ms.assetid: 5b3c8939-793f-116f-11b8-1d4170d95a63
ms.date: 06/08/2017
---


# Application Object (Visio)

Represents an instance of Visio. An external program typically creates or retrieves an  **Application** object before it can retrieve other Visio objects from that instance. Use the Microsoft Visual Basic **CreateObject** function or the **New** keyword to run a new instance, or use the **GetObject** function to retrieve an instance that is already running. You can also use the **CreateObject** function with the **InvisibleApp** object to run a new instance that is invisible. Set the value of the **InvisibleApp** object's **Visible** property to **True** to show it.


## Remarks

Use the  **Documents**, **Windows**, and **Addons** properties of an **Application** object to retrieve the **Document**, **Window**, and **Addon** collections of the instance.

Use the  **ActiveDocument**, **ActivePage**, or **ActiveWindow** property to retrieve the currently active **Document**, **Page**, or **Window** object.


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Use the  **BuiltInMenus**, **BuiltInToolbars**, **CustomMenus**, **CustomToolbars**, or **CommandBars** property to access the **Application** object's menus and toolbars.

 **ActiveDocument** is the default property of an **Application** object.


 **Note**  Code in the Microsoft Visual Basic for Applications project of a Visio document can use the Visio global object instead of a Visio  **Application** object to retrieve other objects.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this object maps to the following types:


-  **Microsoft.Office.Interop.Visio.ApplicationClass** (to access the **Application** object.)
    
-  **Microsoft.Office.Interop.Visio.ApplicationClass.Application** (to construct the **Application** object.)
    
-  **Microsoft.Office.Interop.Visio.EApplication_Event** (to access events on the **Application** object.
    

## Events



|**Name**|
|:-----|
|[AfterModal](http://msdn.microsoft.com/library/e19a0ef3-349c-1d7f-9856-7ef6c66f5f0e%28Office.15%29.aspx)|
|[AfterRemoveHiddenInformation](http://msdn.microsoft.com/library/abd8501a-b528-0433-1633-6d26960dcdaa%28Office.15%29.aspx)|
|[AfterReplaceShapes](http://msdn.microsoft.com/library/b02de031-086a-41cc-d832-5434b8096444%28Office.15%29.aspx)|
|[AfterResume](http://msdn.microsoft.com/library/73cac713-6559-ae7c-32a6-5c421302a3d9%28Office.15%29.aspx)|
|[AfterResumeEvents](http://msdn.microsoft.com/library/c4a662a9-575f-c9db-05b8-d71b4459793b%28Office.15%29.aspx)|
|[AppActivated](http://msdn.microsoft.com/library/150864ab-574a-6556-a56a-8ca619796062%28Office.15%29.aspx)|
|[AppDeactivated](http://msdn.microsoft.com/library/362bb2fb-91a2-01be-e686-3bf076388341%28Office.15%29.aspx)|
|[AppObjActivated](http://msdn.microsoft.com/library/ab27fad1-5afb-534c-987f-e5401603aa52%28Office.15%29.aspx)|
|[AppObjDeactivated](http://msdn.microsoft.com/library/0a401a6e-6aee-3175-6834-55a828a9c864%28Office.15%29.aspx)|
|[BeforeDataRecordsetDelete](http://msdn.microsoft.com/library/b0da57d0-d87f-410c-cfdc-abf8a7bd4b3b%28Office.15%29.aspx)|
|[BeforeDocumentClose](http://msdn.microsoft.com/library/c0d7815e-25bb-7b7e-f80b-81472edc47ca%28Office.15%29.aspx)|
|[BeforeDocumentSave](http://msdn.microsoft.com/library/d5d159fb-52e8-2308-6cc2-3b5b4f82fabb%28Office.15%29.aspx)|
|[BeforeDocumentSaveAs](http://msdn.microsoft.com/library/e6782126-d2e7-c82e-b4dc-a9a5cece14b7%28Office.15%29.aspx)|
|[BeforeMasterDelete](http://msdn.microsoft.com/library/65e3bbed-46f4-25c1-1e3f-af61ef61cce9%28Office.15%29.aspx)|
|[BeforeModal](http://msdn.microsoft.com/library/505d3e54-c8f7-7f02-90d2-43f73573b296%28Office.15%29.aspx)|
|[BeforePageDelete](http://msdn.microsoft.com/library/658e3367-2f5b-e2d4-6c07-9b4463ee500a%28Office.15%29.aspx)|
|[BeforeQuit](http://msdn.microsoft.com/library/ad5ed704-4e7e-f8a9-b238-3c552dc3f292%28Office.15%29.aspx)|
|[BeforeReplaceShapes](http://msdn.microsoft.com/library/fbf44569-0539-9292-ce20-1f9e34238b33%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/4384f7b1-9e88-9a73-a452-5943fb40f18b%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/b33b646c-be39-8f34-d62e-2fcc0283c675%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/db6cdf8c-6a1d-37c4-e185-8809ddafc340%28Office.15%29.aspx)|
|[BeforeStyleDelete](http://msdn.microsoft.com/library/5fc9abed-dc07-0af8-0c3b-87ecabc204a0%28Office.15%29.aspx)|
|[BeforeSuspend](http://msdn.microsoft.com/library/6649fea7-017c-9295-12b5-f350dcf38b28%28Office.15%29.aspx)|
|[BeforeSuspendEvents](http://msdn.microsoft.com/library/a6879424-40d8-e517-aad0-f31aa84a49f6%28Office.15%29.aspx)|
|[BeforeWindowClosed](http://msdn.microsoft.com/library/e062ffe4-8680-456c-4aea-3669e1cab20d%28Office.15%29.aspx)|
|[BeforeWindowPageTurn](http://msdn.microsoft.com/library/ddb79c04-7cb4-61fd-a37d-d5969e445d5c%28Office.15%29.aspx)|
|[BeforeWindowSelDelete](http://msdn.microsoft.com/library/36ff6935-23a8-b155-e5d1-58ae90b10cb6%28Office.15%29.aspx)|
|[CalloutRelationshipAdded](http://msdn.microsoft.com/library/f4ab588e-509d-e11a-4ecd-060c67cbdfe3%28Office.15%29.aspx)|
|[CalloutRelationshipDeleted](http://msdn.microsoft.com/library/779e962c-85f7-e25e-22f7-529b392b93a2%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/8c07be33-8d0d-4957-7f08-daef8b798f28%28Office.15%29.aspx)|
|[ConnectionsAdded](http://msdn.microsoft.com/library/bde55734-25c0-8b8d-231d-a597e99a1d2e%28Office.15%29.aspx)|
|[ConnectionsDeleted](http://msdn.microsoft.com/library/9578be17-8c77-9454-c8a8-1e02fa6516b2%28Office.15%29.aspx)|
|[ContainerRelationshipAdded](http://msdn.microsoft.com/library/8d69056a-9814-d521-86ed-8cdbfa1aeb56%28Office.15%29.aspx)|
|[ContainerRelationshipDeleted](http://msdn.microsoft.com/library/1aa5cd59-f350-ba47-0654-dc1bf1d6073f%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/0cc49837-c819-774c-c69b-45ae86b7fa0d%28Office.15%29.aspx)|
|[DataRecordsetAdded](http://msdn.microsoft.com/library/04a54ec4-6f87-ac4d-f35c-bc3debca4a65%28Office.15%29.aspx)|
|[DataRecordsetChanged](http://msdn.microsoft.com/library/8be61b1a-3a3c-5880-47bc-e2cac9bb60f9%28Office.15%29.aspx)|
|[DesignModeEntered](http://msdn.microsoft.com/library/312f0bda-1375-e176-f5c5-4ebd3c9c8b6d%28Office.15%29.aspx)|
|[DocumentChanged](http://msdn.microsoft.com/library/bed6b530-8d95-10f1-2239-ae7fa940db76%28Office.15%29.aspx)|
|[DocumentCloseCanceled](http://msdn.microsoft.com/library/138e4bf9-87e7-dc9b-4cf6-b12992f22e20%28Office.15%29.aspx)|
|[DocumentCreated](http://msdn.microsoft.com/library/322aaaab-97db-61a7-22f7-65645e1d2f2f%28Office.15%29.aspx)|
|[DocumentOpened](http://msdn.microsoft.com/library/daaf496c-1c9c-cdc1-a06c-ac8cc8fd912f%28Office.15%29.aspx)|
|[DocumentSaved](http://msdn.microsoft.com/library/a11744f6-a1a7-41db-c427-5bae96b9b0ec%28Office.15%29.aspx)|
|[DocumentSavedAs](http://msdn.microsoft.com/library/f03e5fe2-04da-8324-fc0a-be16daf3ad30%28Office.15%29.aspx)|
|[EnterScope](http://msdn.microsoft.com/library/f7935021-2458-cc8e-dd25-d8d2eb16fa6d%28Office.15%29.aspx)|
|[ExitScope](http://msdn.microsoft.com/library/9306972d-6d07-fa82-507d-d4e6d8c80e17%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/f6414b65-cd58-f253-df26-ac33f821799c%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/becaba95-3904-fa18-37a2-b8b8b48a11ab%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/3e9481cc-b9e7-17c0-7b7d-93b6fa2f8825%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/d044400a-e552-6615-ce2c-1d0aec723b6f%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/fb638bc4-8226-de1c-6609-4b757b7d0e4c%28Office.15%29.aspx)|
|[MarkerEvent](http://msdn.microsoft.com/library/1d0c20cc-ccfd-595c-04ea-afce487e582c%28Office.15%29.aspx)|
|[MasterAdded](http://msdn.microsoft.com/library/ef5ddfa4-3f33-e913-ea96-a1b063a1af2b%28Office.15%29.aspx)|
|[MasterChanged](http://msdn.microsoft.com/library/f92d988d-1cbb-00c1-9d5d-46f001e76433%28Office.15%29.aspx)|
|[MasterDeleteCanceled](http://msdn.microsoft.com/library/8dabb35b-8959-ef83-90fd-3287265f60a5%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/f65b3ee7-9b34-d09f-220f-3c7d39a40565%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/3ffd86f8-8700-88a7-9ffc-24df11c93dd4%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/8ca78f5e-5287-0ef5-57ea-d7d116f45ff0%28Office.15%29.aspx)|
|[MustFlushScopeBeginning](http://msdn.microsoft.com/library/98a47603-19c0-4588-3d65-1f9d3fe118c1%28Office.15%29.aspx)|
|[MustFlushScopeEnded](http://msdn.microsoft.com/library/ba9ae16a-9cc6-79d6-d838-e5927937c142%28Office.15%29.aspx)|
|[NoEventsPending](http://msdn.microsoft.com/library/8cb93f89-4541-53f8-a95c-abf5b349f67d%28Office.15%29.aspx)|
|[OnKeystrokeMessageForAddon](http://msdn.microsoft.com/library/0b3fcabc-217f-fa68-d139-455286b3a34f%28Office.15%29.aspx)|
|[PageAdded](http://msdn.microsoft.com/library/16813cbf-d4e0-17b1-308e-06e2a3adf0d4%28Office.15%29.aspx)|
|[PageChanged](http://msdn.microsoft.com/library/bcb49753-6980-307f-362d-92cebe7bdf53%28Office.15%29.aspx)|
|[PageDeleteCanceled](http://msdn.microsoft.com/library/1efbd64c-365b-c967-59bb-8314a0fa6f34%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/1b47836b-def8-6019-93f5-1694fd7cb4f9%28Office.15%29.aspx)|
|[QueryCancelDocumentClose](http://msdn.microsoft.com/library/5d58168d-ed84-943e-26b6-16246c907e52%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/b22d2387-4586-fb6d-0cfe-83088f807a47%28Office.15%29.aspx)|
|[QueryCancelMasterDelete](http://msdn.microsoft.com/library/8277a799-c86f-ddd4-7c0a-da0762418217%28Office.15%29.aspx)|
|[QueryCancelPageDelete](http://msdn.microsoft.com/library/81e9ab8a-5060-9ebf-52c7-e22ed45487f1%28Office.15%29.aspx)|
|[QueryCancelQuit](http://msdn.microsoft.com/library/19b58edc-dafd-acad-deee-19b2b4021ab6%28Office.15%29.aspx)|
|[QueryCancelReplaceShapes](http://msdn.microsoft.com/library/50c0f2a6-f534-f3af-7e83-c865abda8bf9%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/dc1c6b8a-1c60-06fb-9c8f-2919d0081838%28Office.15%29.aspx)|
|[QueryCancelStyleDelete](http://msdn.microsoft.com/library/7f3ce781-67d8-7a6e-d8f0-b077c8956b12%28Office.15%29.aspx)|
|[QueryCancelSuspend](http://msdn.microsoft.com/library/1beb9459-f331-d20b-59f0-da505a375a4f%28Office.15%29.aspx)|
|[QueryCancelSuspendEvents](http://msdn.microsoft.com/library/886fa424-67b3-6a4d-f0bb-99ee646b0753%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/67d3b9e1-c2f3-20ba-0bb5-3ab2dc8f1564%28Office.15%29.aspx)|
|[QueryCancelWindowClose](http://msdn.microsoft.com/library/f4ac803c-5a65-a310-f731-1d2666638525%28Office.15%29.aspx)|
|[QuitCanceled](http://msdn.microsoft.com/library/0861a2ea-f4d7-dc57-7642-2e7642fd2afe%28Office.15%29.aspx)|
|[ReplaceShapesCanceled](http://msdn.microsoft.com/library/e8eecd64-e4bd-d2c4-b942-c5ff607a4121%28Office.15%29.aspx)|
|[RuleSetValidated](http://msdn.microsoft.com/library/d074d4d9-9840-0054-8502-e8537952d7d0%28Office.15%29.aspx)|
|[RunModeEntered](http://msdn.microsoft.com/library/3a8827d9-ff0c-a1c4-2848-72758277aff4%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/76a6c2c6-c2ab-97ad-a906-0780a81f7eb2%28Office.15%29.aspx)|
|[SelectionChanged](http://msdn.microsoft.com/library/d2749204-9003-f4a7-1de0-b47d5e6abb1b%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/612b087f-1985-f399-44ad-7308344ae97f%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/55024b4a-44f1-512e-7739-d1258960e988%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/aac5dfc5-370e-8299-4e3e-39fe9a7000d2%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/2b08879a-9607-c878-9524-6806e43e08ae%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/a7f04e35-9d36-69fa-637f-4930604037f1%28Office.15%29.aspx)|
|[ShapeLinkAdded](http://msdn.microsoft.com/library/24b517f7-e6da-df93-db2e-14740050f832%28Office.15%29.aspx)|
|[ShapeLinkDeleted](http://msdn.microsoft.com/library/c1ae3fda-d5fb-210e-7e84-98ffde8bbd29%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/321f937c-27e0-be80-9d6a-78e4e85629ec%28Office.15%29.aspx)|
|[StyleAdded](http://msdn.microsoft.com/library/a966cbc6-529e-5525-5fc2-ed9cd3250dfa%28Office.15%29.aspx)|
|[StyleChanged](http://msdn.microsoft.com/library/f56680b3-71c3-91c6-23d0-7d5840f9aeb5%28Office.15%29.aspx)|
|[StyleDeleteCanceled](http://msdn.microsoft.com/library/c5d2960f-1fd2-0371-93c0-566ab541dc97%28Office.15%29.aspx)|
|[SuspendCanceled](http://msdn.microsoft.com/library/63b2a2c6-5ac7-2e04-e7ac-3295df179498%28Office.15%29.aspx)|
|[SuspendEventsCanceled](http://msdn.microsoft.com/library/33892ba1-90b2-30ee-d355-e3c353749ea8%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/b1d5b023-11ba-193f-e5ab-807940f6d84d%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/2b1ed000-b755-913e-b531-cc6a5a224ac4%28Office.15%29.aspx)|
|[ViewChanged](http://msdn.microsoft.com/library/2cb8dbfb-100c-1fe3-05c5-bb9a2d97075a%28Office.15%29.aspx)|
|[VisioIsIdle](http://msdn.microsoft.com/library/58a66628-d8df-f55c-7d25-e6b272b37906%28Office.15%29.aspx)|
|[WindowActivated](http://msdn.microsoft.com/library/ef89f592-b457-b170-0e2e-84d9e1c572f2%28Office.15%29.aspx)|
|[WindowChanged](http://msdn.microsoft.com/library/29bb6ea8-4558-38c4-941f-839cd119abba%28Office.15%29.aspx)|
|[WindowCloseCanceled](http://msdn.microsoft.com/library/1273b75d-0543-69aa-aab3-47281295ee6b%28Office.15%29.aspx)|
|[WindowOpened](http://msdn.microsoft.com/library/a75a50b5-9784-e191-991a-ca9b41994ff9%28Office.15%29.aspx)|
|[WindowTurnedToPage](http://msdn.microsoft.com/library/f747ed48-6da1-fd7f-4cdd-e9f46f02b1d0%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddUndoUnit](http://msdn.microsoft.com/library/90542078-5efa-fec6-b853-41f8a998bea9%28Office.15%29.aspx)|
|[BeginUndoScope](http://msdn.microsoft.com/library/7e3a4e34-6796-4277-1dc4-7252ee2b6720%28Office.15%29.aspx)|
|[ClearCustomMenus](http://msdn.microsoft.com/library/01c7f266-e940-b02c-b77d-7178c9296f98%28Office.15%29.aspx)|
|[ClearCustomToolbars](http://msdn.microsoft.com/library/fa9ad39a-2765-b172-a7ad-140f9bb845b9%28Office.15%29.aspx)|
|[ConvertResult](http://msdn.microsoft.com/library/b326c9cf-a7f3-33d7-1b29-8d1360301a9d%28Office.15%29.aspx)|
|[DoCmd](http://msdn.microsoft.com/library/096c11a0-1234-6a47-7bc4-5f808acfe21a%28Office.15%29.aspx)|
|[EndUndoScope](http://msdn.microsoft.com/library/352188d2-8a2a-1a6d-674e-93fce9245810%28Office.15%29.aspx)|
|[EnumDirectories](http://msdn.microsoft.com/library/71ed7f7f-3428-5c50-2ab9-5452188dcfe0%28Office.15%29.aspx)|
|[FormatResult](http://msdn.microsoft.com/library/1b2178ab-e2ed-b618-ad2a-d18196f50be2%28Office.15%29.aspx)|
|[FormatResultEx](http://msdn.microsoft.com/library/68dadf46-0d2b-2a2d-a389-0a17c84e45b4%28Office.15%29.aspx)|
|[GetBuiltInStencilFile](http://msdn.microsoft.com/library/2ae65aaa-d441-c7e8-3c8c-737bcca84738%28Office.15%29.aspx)|
|[GetCustomStencilFile](http://msdn.microsoft.com/library/10c8ec1d-f4e0-07dd-4487-40f85cbf5497%28Office.15%29.aspx)|
|[GetPreviewEnabled](http://msdn.microsoft.com/library/6e0d42b9-f1d4-d8b9-ab9c-7f00ba6c6a9d%28Office.15%29.aspx)|
|[InvokeHelp](http://msdn.microsoft.com/library/dffc0412-9b90-466c-c0f9-d32f702d4927%28Office.15%29.aspx)|
|[OnComponentEnterState](http://msdn.microsoft.com/library/f5d61cb0-d7c0-df13-f7c4-b39c7104f73a%28Office.15%29.aspx)|
|[PurgeUndo](http://msdn.microsoft.com/library/d5d18607-2b1d-6b47-2a81-43345ff0be8a%28Office.15%29.aspx)|
|[QueueMarkerEvent](http://msdn.microsoft.com/library/2afa9553-db06-12ca-f5ef-28431f56a92d%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/1f8b73cd-10bd-e571-eee4-db05d9aa12cc%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/ab7ac8bc-e747-9188-1546-6bb31f77231b%28Office.15%29.aspx)|
|[RegisterRibbonX](http://msdn.microsoft.com/library/178db1c3-b3af-aa3f-af03-1aec1eab549a%28Office.15%29.aspx)|
|[RenameCurrentScope](http://msdn.microsoft.com/library/0ccd9c6b-704c-b956-5ea9-4f1ed01baee7%28Office.15%29.aspx)|
|[SetCustomMenus](http://msdn.microsoft.com/library/90aa627c-ba51-87a7-4347-6a806998e1a4%28Office.15%29.aspx)|
|[SetCustomToolbars](http://msdn.microsoft.com/library/fe5a3e40-83ea-d02f-03cd-d0ad758aa408%28Office.15%29.aspx)|
|[SetPreviewEnabled](http://msdn.microsoft.com/library/fa66a148-2eca-85b8-b780-ff077b14d0f2%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/728d9af0-c9f2-c3ff-5ed3-a20e8a507a6a%28Office.15%29.aspx)|
|[UnregisterRibbonX](http://msdn.microsoft.com/library/83c5a7c3-b3bb-cbbd-6857-2ae822e599f6%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Active](http://msdn.microsoft.com/library/d2e8e683-15b8-9c6e-f945-5a1d17a177b0%28Office.15%29.aspx)|
|[ActiveDocument](http://msdn.microsoft.com/library/0dbc95b6-4920-4291-55c0-ec0128e7f006%28Office.15%29.aspx)|
|[ActivePage](http://msdn.microsoft.com/library/1d0496aa-a6f5-0886-fb8f-8071f95fa333%28Office.15%29.aspx)|
|[ActivePrinter](http://msdn.microsoft.com/library/1b0587d1-75e0-3a1d-963c-f4fb29e52d8c%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/6da310fd-3fb1-618b-d80f-98ee1e45d5a2%28Office.15%29.aspx)|
|[AddonPaths](http://msdn.microsoft.com/library/ec6cf92d-5570-8c24-87c2-68f26f3721a4%28Office.15%29.aspx)|
|[Addons](http://msdn.microsoft.com/library/c0d9731e-124f-b308-4c84-a14e0b82ff00%28Office.15%29.aspx)|
|[AlertResponse](http://msdn.microsoft.com/library/aa7a14b1-b2df-daa6-7298-66880a573f5c%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/be058d51-6bfa-c653-da44-fa38e0b96c63%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/d2ac6782-7b80-8760-b7a1-27503182c85a%28Office.15%29.aspx)|
|[AutoLayout](http://msdn.microsoft.com/library/b631def8-d271-8ed0-880a-db8a1ee26759%28Office.15%29.aspx)|
|[AutoRecoverInterval](http://msdn.microsoft.com/library/06aa731b-b426-a1a2-a25b-8ac32133eb1a%28Office.15%29.aspx)|
|[AvailablePrinters](http://msdn.microsoft.com/library/bd070ee3-4f32-1ff0-427c-d61b7778e6c5%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/92fcdbe9-dfb1-cd20-4700-796bf7ca17f1%28Office.15%29.aspx)|
|[BuiltInMenus](http://msdn.microsoft.com/library/0f76537c-5d9b-bcfa-c528-4644bd0375d5%28Office.15%29.aspx)|
|[BuiltInToolbars](http://msdn.microsoft.com/library/e0460fa5-23da-f452-f541-feabe8e3bffb%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/182ea1e1-f896-f619-1bf0-df4a57b39abf%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/3829b033-aed4-a132-ff44-96d419dd09cd%28Office.15%29.aspx)|
|[CommandLine](http://msdn.microsoft.com/library/36c8320e-17b4-111d-1b2c-e8f7096e1680%28Office.15%29.aspx)|
|[ConnectorToolDataObject](http://msdn.microsoft.com/library/7b1eedad-3d62-c2a1-5ba7-200a594ba32f%28Office.15%29.aspx)|
|[CurrentEdition](http://msdn.microsoft.com/library/11484259-abd3-d727-ff2e-b9bc07fe9c5a%28Office.15%29.aspx)|
|[CurrentScope](http://msdn.microsoft.com/library/a45fd841-efb4-90b6-65fb-21f9f8e8ea0c%28Office.15%29.aspx)|
|[CustomMenus](http://msdn.microsoft.com/library/c8ccb1fa-2654-17db-166f-c724da345626%28Office.15%29.aspx)|
|[CustomMenusFile](http://msdn.microsoft.com/library/88a3f298-36a4-892d-33fc-8fe330d51437%28Office.15%29.aspx)|
|[CustomToolbars](http://msdn.microsoft.com/library/1c945955-af48-5dd1-f186-d7d0cf02e6d2%28Office.15%29.aspx)|
|[CustomToolbarsFile](http://msdn.microsoft.com/library/e4759ee0-1128-8238-ad0b-47ad365ce88d%28Office.15%29.aspx)|
|[DataFeaturesEnabled](http://msdn.microsoft.com/library/3ff6eb4e-1ea8-3f53-c86b-133d4960516e%28Office.15%29.aspx)|
|[DefaultAngleUnits](http://msdn.microsoft.com/library/28c51825-bff1-8fca-2070-76912593c53b%28Office.15%29.aspx)|
|[DefaultDurationUnits](http://msdn.microsoft.com/library/11810de2-0c2f-a498-6b7a-090d5397066b%28Office.15%29.aspx)|
|[DefaultRectangleDataObject](http://msdn.microsoft.com/library/22e7f5ff-516d-4bd0-82bf-2363d1cad973%28Office.15%29.aspx)|
|[DefaultTextUnits](http://msdn.microsoft.com/library/54d2ce66-a8e9-b45e-0ec1-f0e7e39e8c5a%28Office.15%29.aspx)|
|[DefaultZoomBehavior](http://msdn.microsoft.com/library/59f26e36-90e3-defa-be04-b7a8ce710eeb%28Office.15%29.aspx)|
|[DeferRecalc](http://msdn.microsoft.com/library/25f7ee2e-8987-f03e-5dee-74550bc19f83%28Office.15%29.aspx)|
|[DeferRelationshipRecalc](http://msdn.microsoft.com/library/b85ce4e4-4425-e508-042f-4119353a60b8%28Office.15%29.aspx)|
|[DialogFont](http://msdn.microsoft.com/library/8742b97f-7f66-38c7-fafd-a343c1160671%28Office.15%29.aspx)|
|[Documents](http://msdn.microsoft.com/library/dee2a72f-526c-7b10-57b4-c4fbca43b083%28Office.15%29.aspx)|
|[DrawingPaths](http://msdn.microsoft.com/library/46a0a769-8ef4-1cc3-9206-24c23b0982ee%28Office.15%29.aspx)|
|[EventInfo](http://msdn.microsoft.com/library/19065ecc-62bb-5bc4-fdfa-452ab6224211%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/1c72aac3-1714-8d00-831c-e049572de1eb%28Office.15%29.aspx)|
|[EventsEnabled](http://msdn.microsoft.com/library/92775825-c17d-fd4f-a315-7a181d75aed5%28Office.15%29.aspx)|
|[FullBuild](http://msdn.microsoft.com/library/608b99df-027b-7878-e519-311b57dc86bd%28Office.15%29.aspx)|
|[HelpPaths](http://msdn.microsoft.com/library/eba05b64-61d8-970d-65f4-26ea41840fcf%28Office.15%29.aspx)|
|[InhibitSelectChange](http://msdn.microsoft.com/library/d3673adf-a8e2-bc85-aa56-232ec3a93588%28Office.15%29.aspx)|
|[InstanceHandle32](http://msdn.microsoft.com/library/d9e51540-21d5-5f52-68ef-1d49cb30cd51%28Office.15%29.aspx)|
|[InstanceHandle64](http://msdn.microsoft.com/library/213b7c36-b443-2b1b-7f2c-851747d03fff%28Office.15%29.aspx)|
|[IsInScope](http://msdn.microsoft.com/library/adb9a52f-8e62-9d92-d8bf-81bed48b2cc3%28Office.15%29.aspx)|
|[IsUndoingOrRedoing](http://msdn.microsoft.com/library/c398cff2-90df-7a7e-b810-fdda8cbfbe8a%28Office.15%29.aspx)|
|[IsVisio32](http://msdn.microsoft.com/library/14dc8f6b-3548-f76e-50da-cb19426b171f%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/78dc3295-16bd-28fd-43d7-4e6d7924e3be%28Office.15%29.aspx)|
|[LanguageHelp](http://msdn.microsoft.com/library/71ae2f5a-5a8c-ea38-e9db-081bc8fe5cc4%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/3fa0c4a4-3a1c-b035-9f9d-e4358917ebee%28Office.15%29.aspx)|
|[LiveDynamics](http://msdn.microsoft.com/library/fc5a887b-318a-fd25-c2b5-52d6cc1c026e%28Office.15%29.aspx)|
|[MyShapesPath](http://msdn.microsoft.com/library/0e5a598b-262f-dbf0-3c68-3199750fd5a9%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/d30a1b28-7ef8-e77b-220c-16eb9b6f8562%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/50baf864-034e-9051-3671-a3c8f0f368ed%28Office.15%29.aspx)|
|[OnDataChangeDelay](http://msdn.microsoft.com/library/14952e41-445a-77ff-30f7-e7aa6d8fcc32%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/ac19f086-3a14-64b0-6ecf-94ba7ac54cf5%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/455474f3-f39f-cc4c-4e6a-e6dd907c2b35%28Office.15%29.aspx)|
|[ProcessID](http://msdn.microsoft.com/library/d089bfa9-83a4-1b44-80ab-f23c5198801f%28Office.15%29.aspx)|
|[PromptForSummary](http://msdn.microsoft.com/library/6250acdc-ed15-5d07-cbbe-8a4b400d775d%28Office.15%29.aspx)|
|[SaveAsWebObject](http://msdn.microsoft.com/library/ce3f8cb0-8e22-e364-e7d8-b1fc3506bc59%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/934e697f-da6c-5793-433b-dddb5d806920%28Office.15%29.aspx)|
|[Settings](http://msdn.microsoft.com/library/b62413cb-a038-2679-8701-47ba700a93c4%28Office.15%29.aspx)|
|[ShowChanges](http://msdn.microsoft.com/library/6a8a7ee7-ad57-1d52-8a22-fb30be076236%28Office.15%29.aspx)|
|[ShowProgress](http://msdn.microsoft.com/library/4dcfcec7-d652-0b52-a4e8-a43122e72988%28Office.15%29.aspx)|
|[ShowStatusBar](http://msdn.microsoft.com/library/a6eade7f-b056-92ef-0a57-acd466f6a99a%28Office.15%29.aspx)|
|[ShowToolbar](http://msdn.microsoft.com/library/274dbfae-30bd-02cb-c8c4-246a3a3f26ef%28Office.15%29.aspx)|
|[StartupPaths](http://msdn.microsoft.com/library/966a91d9-9ada-d0e1-9886-271ea47faaf9%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/59199a84-6272-e160-429b-0c9c32dc4f91%28Office.15%29.aspx)|
|[StencilPaths](http://msdn.microsoft.com/library/1b664a6d-ba52-7115-7c48-bf2f6dd8068d%28Office.15%29.aspx)|
|[TemplatePaths](http://msdn.microsoft.com/library/149a9ef2-e255-3dad-2177-b29c173fa66d%28Office.15%29.aspx)|
|[TraceFlags](http://msdn.microsoft.com/library/aae7879a-7f21-f16e-cfcc-2520c70af7e7%28Office.15%29.aspx)|
|[TypelibMajorVersion](http://msdn.microsoft.com/library/17e1abf3-5a5d-aac9-9f78-4eeb2c4a6c79%28Office.15%29.aspx)|
|[TypelibMinorVersion](http://msdn.microsoft.com/library/ee3a31db-ddfe-a036-a570-43e6f27ad024%28Office.15%29.aspx)|
|[UndoEnabled](http://msdn.microsoft.com/library/54890621-84c3-8bde-2043-acb91a5b85dc%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/2f122cb1-735f-ceb8-76db-5b7a80bce080%28Office.15%29.aspx)|
|[VBAEnabled](http://msdn.microsoft.com/library/fd4aa300-2117-aa66-54da-3be7be920287%28Office.15%29.aspx)|
|[Vbe](http://msdn.microsoft.com/library/1ad29679-1078-5682-e375-868e32fb0ca5%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/c2e3b022-507d-c73c-6fa4-9689cc5600f3%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/48b0a402-3655-b6aa-19da-145d2c7ceadf%28Office.15%29.aspx)|
|[Window](http://msdn.microsoft.com/library/fd996e7d-a204-ab0d-538a-439c61532bb9%28Office.15%29.aspx)|
|[WindowHandle32](http://msdn.microsoft.com/library/d4c653ae-6582-0d86-75ee-969fe978e754%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/d8924555-fbe8-b423-523b-958d50955c37%28Office.15%29.aspx)|

