---
title: Shape Object (Visio)
keywords: vis_sdr.chm10225
f1_keywords:
- vis_sdr.chm10225
ms.prod: visio
api_name:
- Visio.Shape
ms.assetid: da7a8872-4ebb-a607-e0ed-eebf68ff5630
ms.date: 06/08/2017
---


# Shape Object (Visio)

Represents anything you can select in a drawing window: a basic shape, a group, a guide, or an object from another application embedded or linked in Microsoft Visio.


## Remarks

The default property of a  **Shape** object is **Name**.

You can retrieve a particular  **Shape** object from the **Shapes** collection of the following objects:




-  **Page** object
    
-  **Master** object
    
-  **Shape** object that represents a group
    


To retrieve  **Cell** objects and **Connect** objects, use the **Cells** and **Connects** properties of a **Shape** object, respectively.


 **Note**  The **PageSheet** property of a **Page** object and **Master** object returns a **Shape** object whose **Type** property returns **visTypePage**. It has cells that specify properties such as drawing size and drawing scale. The **DocumentSheet** property of a **Document** object also returns a **Shape** object whose **Type** property returns **visTypeDoc**. It has cells that specify properties of the document.If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this object maps to the following types:


## Events



|**Name**|
|:-----|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/3979ee0b-155d-7c16-8141-b2131270b6c6%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/6cbfc832-cdf6-1289-feb4-1b1fcbb3574f%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/f64b57b6-c92c-dd17-9698-211d9ca2fe83%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/d3324bb1-f944-e644-79ef-55022b31fbd2%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/f5b312cf-97ab-15c8-3d1c-07edd2023a40%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/cf141b03-5eaf-bf42-a64f-049f8dec2a14%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/89ce290b-a164-4581-b83d-64d205765aeb%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/18fffdd9-2d6a-90d5-ac34-ce6f3a5e8df6%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/a2283176-3584-317e-3645-9e6f3dece076%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/d050cf74-b427-32ef-fe11-77246bb9cf55%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/de7ffc8b-ad3d-2738-4470-be9d34c43b69%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/ca63a476-a7d0-bd27-6c41-5e36b4ef56ed%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/10811705-9619-d4d8-80f5-f1fa08eed52f%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/89e562f4-f3b0-54bd-cbac-515eecb70c97%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/3c31acbc-99c9-f047-7aaa-01eddf4242ea%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/6c4a9bab-cad0-5f37-a1f8-ca040526e1b5%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/ba707fd6-2a5a-65f6-6db4-ed3b5250a103%28Office.15%29.aspx)|
|[ShapeLinkAdded](http://msdn.microsoft.com/library/5cd7431f-18da-184c-7976-06f174cd5f73%28Office.15%29.aspx)|
|[ShapeLinkDeleted](http://msdn.microsoft.com/library/9233b720-f228-0403-d705-15f5eb39e3b4%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/b26b9740-a3bf-1100-0f7b-f34cb03be53c%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/e6516896-de9e-e90f-679b-541c15ab26db%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/aca15d4f-c623-471b-80b2-80f6afd2d5c7%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddGuide](http://msdn.microsoft.com/library/1155354e-3855-4def-bafb-0d70c933a57a%28Office.15%29.aspx)|
|[AddHyperlink](http://msdn.microsoft.com/library/fbf77a65-88a1-e710-60a2-efde9e7df968%28Office.15%29.aspx)|
|[AddNamedRow](http://msdn.microsoft.com/library/c18380b1-418d-454f-3c90-fa4624291628%28Office.15%29.aspx)|
|[AddRow](http://msdn.microsoft.com/library/8b8dcf65-9b42-b3bf-0da3-61d3fbd02996%28Office.15%29.aspx)|
|[AddRows](http://msdn.microsoft.com/library/8b267f98-e077-0854-a1aa-a0ce8719a2c5%28Office.15%29.aspx)|
|[AddSection](http://msdn.microsoft.com/library/64396db4-8361-ece9-b029-24d62ba0a290%28Office.15%29.aspx)|
|[AddToContainers](http://msdn.microsoft.com/library/ddd3f532-cbbf-3c63-0e02-49f4ea8ca90c%28Office.15%29.aspx)|
|[AutoConnect](http://msdn.microsoft.com/library/36b634be-9943-1aec-f8e0-70467b82eed1%28Office.15%29.aspx)|
|[BoundingBox](http://msdn.microsoft.com/library/68053d27-b7da-9ae7-7557-5d49c8d3e1d6%28Office.15%29.aspx)|
|[BreakLinkToData](http://msdn.microsoft.com/library/1f4ed559-061e-f016-739c-e760e634dba8%28Office.15%29.aspx)|
|[BringForward](http://msdn.microsoft.com/library/88e5c746-e7f2-eddd-35c9-2d9c09c1a602%28Office.15%29.aspx)|
|[BringToFront](http://msdn.microsoft.com/library/91689605-16b4-eda5-2513-3e04f78fc13e%28Office.15%29.aspx)|
|[CenterDrawing](http://msdn.microsoft.com/library/2ac35f58-2f9d-4139-6477-7e865713c863%28Office.15%29.aspx)|
|[ChangePicture](http://msdn.microsoft.com/library/9193d802-cebd-2bfd-5f8e-400fac36c1a5%28Office.15%29.aspx)|
|[ConnectedShapes](http://msdn.microsoft.com/library/7f5a0ac9-d0a7-d9fe-9ecb-8e8070ab5951%28Office.15%29.aspx)|
|[ConvertToGroup](http://msdn.microsoft.com/library/080db7d0-4283-f8d0-eeca-a6495e6f0536%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/2579682b-1dd3-7579-271d-a9994b91a933%28Office.15%29.aspx)|
|[CreateSelection](http://msdn.microsoft.com/library/205efec7-afa7-87e8-9c31-22395b283496%28Office.15%29.aspx)|
|[CreateSubProcess](http://msdn.microsoft.com/library/efb26247-777f-d468-a8e6-19a9e9c4a343%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/fda7a58c-233b-5864-880e-cfa17f20c175%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/0960d9e1-b091-ea8c-0724-e10a68d8821a%28Office.15%29.aspx)|
|[DeleteEx](http://msdn.microsoft.com/library/df4c164d-576a-acce-3322-7f166eb81e4f%28Office.15%29.aspx)|
|[DeleteRow](http://msdn.microsoft.com/library/892ca523-679d-c707-4aba-e43c011cb718%28Office.15%29.aspx)|
|[DeleteSection](http://msdn.microsoft.com/library/e07981f3-5efe-f4ad-0517-1af4913c3f70%28Office.15%29.aspx)|
|[Disconnect](http://msdn.microsoft.com/library/ece61baa-dfe7-7b61-5c45-49de4cf0e394%28Office.15%29.aspx)|
|[DrawArcByThreePoints](http://msdn.microsoft.com/library/9c00cca4-548e-8f15-1747-897fa5482340%28Office.15%29.aspx)|
|[DrawBezier](http://msdn.microsoft.com/library/d38b56a5-2366-e418-206f-db39bd8e2c82%28Office.15%29.aspx)|
|[DrawCircularArc](http://msdn.microsoft.com/library/538ee927-c34a-c697-8bf1-f134355e6060%28Office.15%29.aspx)|
|[DrawLine](http://msdn.microsoft.com/library/8793104a-0ded-e2ca-54e8-acf987b9c797%28Office.15%29.aspx)|
|[DrawNURBS](http://msdn.microsoft.com/library/e1209142-3902-3231-a019-f6e091978847%28Office.15%29.aspx)|
|[DrawOval](http://msdn.microsoft.com/library/7f561251-251e-6aa9-5291-5919ccce1a9e%28Office.15%29.aspx)|
|[DrawPolyline](http://msdn.microsoft.com/library/79bd8e58-097e-6af3-cc52-435acbeececd%28Office.15%29.aspx)|
|[DrawQuarterArc](http://msdn.microsoft.com/library/7bc281ea-eac8-cdb6-ac4b-c71c93a81827%28Office.15%29.aspx)|
|[DrawRectangle](http://msdn.microsoft.com/library/2d02da32-0938-b019-0fa0-c4ef07d1a318%28Office.15%29.aspx)|
|[DrawSpline](http://msdn.microsoft.com/library/02a66d00-2309-b508-7867-90980addb309%28Office.15%29.aspx)|
|[Drop](http://msdn.microsoft.com/library/bce5f9d1-8684-0ff5-a4a3-3c2adb972310%28Office.15%29.aspx)|
|[DropMany](http://msdn.microsoft.com/library/def60b35-ce19-ec34-9654-355856e26b37%28Office.15%29.aspx)|
|[DropManyU](http://msdn.microsoft.com/library/b3e18874-bb90-334f-e633-3e20133242a1%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/a45fd247-e4ad-8149-3656-af9588f076ef%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/f4051560-8719-ea9c-30eb-33230c95786c%28Office.15%29.aspx)|
|[FitCurve](http://msdn.microsoft.com/library/9055ee19-a021-35b5-1993-6f00c8a5f859%28Office.15%29.aspx)|
|[FlipHorizontal](http://msdn.microsoft.com/library/a1f308a7-1f00-9432-ea26-bc1d784b8451%28Office.15%29.aspx)|
|[FlipVertical](http://msdn.microsoft.com/library/d83d29fb-4292-61c3-b2b4-ba6aed6fe7ad%28Office.15%29.aspx)|
|[GetCustomPropertiesLinkedToData](http://msdn.microsoft.com/library/8a0d783d-f5ee-d6c0-adbd-377cbe65e5f5%28Office.15%29.aspx)|
|[GetCustomPropertyLinkedColumn](http://msdn.microsoft.com/library/0d6e3577-d918-1d33-135a-37a3f09f3eaa%28Office.15%29.aspx)|
|[GetFormulas](http://msdn.microsoft.com/library/51ff9731-802c-2001-c5e6-6f7aeb9d6839%28Office.15%29.aspx)|
|[GetFormulasU](http://msdn.microsoft.com/library/f478abfa-d576-fcb2-5126-464b874355a0%28Office.15%29.aspx)|
|[GetLinkedDataRecordsetIDs](http://msdn.microsoft.com/library/1ce55d6c-02ae-8d5d-f581-b368e830bcf5%28Office.15%29.aspx)|
|[GetLinkedDataRow](http://msdn.microsoft.com/library/55e578a5-da95-9a5c-3d1d-5cc5edeb57a7%28Office.15%29.aspx)|
|[GetResults](http://msdn.microsoft.com/library/7380f8b4-ec22-2271-f4ce-19869264ec25%28Office.15%29.aspx)|
|[GluedShapes](http://msdn.microsoft.com/library/0c9c551d-ce28-f7c6-4656-8120962e8d34%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/fe19f27f-47ad-93ef-1d82-4010d8cb6e47%28Office.15%29.aspx)|
|[HasCategory](http://msdn.microsoft.com/library/91115794-31ab-73b1-d1ec-ca249a57a61f%28Office.15%29.aspx)|
|[HitTest](http://msdn.microsoft.com/library/1250ac1d-32f8-d078-3a01-6e2ce045d254%28Office.15%29.aspx)|
|[Import](http://msdn.microsoft.com/library/07c858ee-0bbc-5ac1-37be-1e853cdf2361%28Office.15%29.aspx)|
|[InsertFromFile](http://msdn.microsoft.com/library/894f69fc-65a7-d0a8-a2ae-e56a73843bc2%28Office.15%29.aspx)|
|[InsertObject](http://msdn.microsoft.com/library/7abc6e96-6822-7237-b695-36f297a076fc%28Office.15%29.aspx)|
|[IsCustomPropertyLinked](http://msdn.microsoft.com/library/e75b910f-fb21-3e39-2ca3-ac0913adccf0%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/f70dfdbb-6501-b9b7-3444-7fa35e98637e%28Office.15%29.aspx)|
|[LinkToData](http://msdn.microsoft.com/library/75dd1543-e643-0c7d-a89a-f0dd09d6d323%28Office.15%29.aspx)|
|[MoveToSubprocess](http://msdn.microsoft.com/library/3688c9d5-5b28-23d7-369a-332649267ffe%28Office.15%29.aspx)|
|[Offset](http://msdn.microsoft.com/library/0a82ed87-cc11-77d3-4337-2694f8431a79%28Office.15%29.aspx)|
|[OpenDrawWindow](http://msdn.microsoft.com/library/5e4106a3-ba72-9e3c-1189-9587d39edd00%28Office.15%29.aspx)|
|[OpenSheetWindow](http://msdn.microsoft.com/library/744b72f5-381a-48fc-407f-20ffe815c54e%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/ce5892be-b5e7-2ca0-7ee1-aa4e602641a2%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/0e3a1006-1664-3b60-5d75-d7d4f77d364d%28Office.15%29.aspx)|
|[RemoveFromContainers](http://msdn.microsoft.com/library/b9dbf604-01f0-675a-a0e1-7b30841ec5c5%28Office.15%29.aspx)|
|[ReplaceShape](http://msdn.microsoft.com/library/b330a63d-4e3f-0c4d-c38c-6ee806670225%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/ce8e9253-e1bb-e542-30eb-f9ac2e4305da%28Office.15%29.aspx)|
|[ReverseEnds](http://msdn.microsoft.com/library/f2e450fa-0f86-6c90-cf58-8ee57f055577%28Office.15%29.aspx)|
|[Rotate90](http://msdn.microsoft.com/library/1c7d526e-f053-d9f5-232a-61cf12de2c6e%28Office.15%29.aspx)|
|[SendBackward](http://msdn.microsoft.com/library/9e43cfd9-c2d3-9042-46e3-39e209ae54aa%28Office.15%29.aspx)|
|[SendToBack](http://msdn.microsoft.com/library/faa9cab5-0b2f-8331-e0df-8b4f4be1e69f%28Office.15%29.aspx)|
|[SetBegin](http://msdn.microsoft.com/library/257a6ec4-b9c4-4c42-3c57-6e53c1d4d526%28Office.15%29.aspx)|
|[SetCenter](http://msdn.microsoft.com/library/9a3c0597-c255-44ab-9268-938acd3c5a69%28Office.15%29.aspx)|
|[SetEnd](http://msdn.microsoft.com/library/5f2c7b85-52b3-9147-a989-b2dce61c3493%28Office.15%29.aspx)|
|[SetFormulas](http://msdn.microsoft.com/library/b2371ff1-4742-e178-3606-133c9c8a1937%28Office.15%29.aspx)|
|[SetQuickStyle](http://msdn.microsoft.com/library/aebe80cb-fae9-0be7-e903-882f6eb58b63%28Office.15%29.aspx)|
|[SetResults](http://msdn.microsoft.com/library/b5dccaf0-776a-3f0c-ca45-2efff3d3f95b%28Office.15%29.aspx)|
|[SwapEnds](http://msdn.microsoft.com/library/54096674-eb4f-4f07-a1cf-701faf3b5fae%28Office.15%29.aspx)|
|[TransformXYFrom](http://msdn.microsoft.com/library/4676e464-83c7-7ff6-e742-becc41436259%28Office.15%29.aspx)|
|[TransformXYTo](http://msdn.microsoft.com/library/dc85cf08-0d83-34ff-8389-94a0f5f05c5e%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/a4ff17b9-6bad-aaf4-ce00-2b529c73f48b%28Office.15%29.aspx)|
|[UpdateAlignmentBox](http://msdn.microsoft.com/library/7076ee5f-f536-77ec-a1f7-518195e3e897%28Office.15%29.aspx)|
|[VisualBoundingBox](http://msdn.microsoft.com/library/6a7d4622-8ba5-c598-4aaa-c6297cb4c008%28Office.15%29.aspx)|
|[XYFromPage](http://msdn.microsoft.com/library/85b04e0b-04e1-a5b5-f6ff-393c57751946%28Office.15%29.aspx)|
|[XYToPage](http://msdn.microsoft.com/library/4a230d63-57a8-3b69-6425-2dca6a2014eb%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/01ad1b62-5a69-9c70-3735-f678a6fa537d%28Office.15%29.aspx)|
|[AreaIU](http://msdn.microsoft.com/library/a9982cd2-9a91-f5e5-7297-360b6d9a1f29%28Office.15%29.aspx)|
|[CalloutsAssociated](http://msdn.microsoft.com/library/c1e32bb2-c946-3919-4f6e-064b5be50d6c%28Office.15%29.aspx)|
|[CalloutTarget](http://msdn.microsoft.com/library/4366753a-c8e2-ba85-54fd-9c74cd21d762%28Office.15%29.aspx)|
|[CellExists](http://msdn.microsoft.com/library/479c4d99-0282-3ab0-2e6f-4a17e48adfab%28Office.15%29.aspx)|
|[CellExistsU](http://msdn.microsoft.com/library/da26e913-39c5-7af5-194d-3bb5dca76678%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/2d90b848-ee2c-d69c-e44e-9c30b04bf776%28Office.15%29.aspx)|
|[CellsRowIndex](http://msdn.microsoft.com/library/7415afcb-9d98-5981-bd33-6ca18116470e%28Office.15%29.aspx)|
|[CellsRowIndexU](http://msdn.microsoft.com/library/425fbf08-d44c-2631-7400-55620fd429ee%28Office.15%29.aspx)|
|[CellsSRC](http://msdn.microsoft.com/library/8fb6fd7b-e0ca-c694-3b9d-5390d4192565%28Office.15%29.aspx)|
|[CellsSRCExists](http://msdn.microsoft.com/library/7d614820-2a64-c3ee-b61c-a7c0dcfb90c8%28Office.15%29.aspx)|
|[CellsU](http://msdn.microsoft.com/library/bee20521-6515-8a3b-e861-104f7cc71c26%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/dcb7fa7b-61ff-df09-8128-2d1ef4e17770%28Office.15%29.aspx)|
|[CharCount](http://msdn.microsoft.com/library/2da9c359-d86c-bdf6-3553-01ded11d9208%28Office.15%29.aspx)|
|[ClassID](http://msdn.microsoft.com/library/b3cb2f9c-1247-9799-69f3-5374a112af95%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/498eca91-beb9-b764-0262-a935e5205710%28Office.15%29.aspx)|
|[Connects](http://msdn.microsoft.com/library/9edaac59-f52e-67ee-6e5a-e11572907785%28Office.15%29.aspx)|
|[ContainerProperties](http://msdn.microsoft.com/library/bc375912-f624-dbdc-3b02-2edf3bf5d8a2%28Office.15%29.aspx)|
|[ContainingMaster](http://msdn.microsoft.com/library/ca262f68-472e-3412-f620-ca837c40378c%28Office.15%29.aspx)|
|[ContainingMasterID](http://msdn.microsoft.com/library/e194cd7c-d7c0-2c08-a0df-764398efa447%28Office.15%29.aspx)|
|[ContainingPage](http://msdn.microsoft.com/library/18fe6146-34eb-9369-603b-b3b316aa23d7%28Office.15%29.aspx)|
|[ContainingPageID](http://msdn.microsoft.com/library/fd33d0d6-571d-47b5-28a7-6fa4aa671312%28Office.15%29.aspx)|
|[ContainingShape](http://msdn.microsoft.com/library/b09bc382-de6c-368e-53bd-c8b01fbc0ae1%28Office.15%29.aspx)|
|[Data1](http://msdn.microsoft.com/library/ca9dda75-4ae2-70f0-46bd-ff5afbba84fc%28Office.15%29.aspx)|
|[Data2](http://msdn.microsoft.com/library/59499252-ee61-d158-5ad8-ceece33a8e9e%28Office.15%29.aspx)|
|[Data3](http://msdn.microsoft.com/library/0d02964d-0296-5142-e7c3-e319ea80c224%28Office.15%29.aspx)|
|[DataGraphic](http://msdn.microsoft.com/library/09c804fe-d0ec-ac88-6620-1a41fc8a507a%28Office.15%29.aspx)|
|[DistanceFrom](http://msdn.microsoft.com/library/2df9e60f-b138-4dde-09ca-af4ee2f6a8d0%28Office.15%29.aspx)|
|[DistanceFromPoint](http://msdn.microsoft.com/library/262b5814-3b86-c3eb-9526-96ec73836ad6%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/235e9100-dd91-cb6b-01e6-893b4f7acdd8%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/513838c2-f00e-06e3-f08b-b23295f7f0d3%28Office.15%29.aspx)|
|[FillStyle](http://msdn.microsoft.com/library/f674da21-deac-4636-608c-c26241a7b125%28Office.15%29.aspx)|
|[FillStyleKeepFmt](http://msdn.microsoft.com/library/39fc0329-322e-fd96-2c42-43bdcd170c02%28Office.15%29.aspx)|
|[ForeignData](http://msdn.microsoft.com/library/c7d826fd-b411-3403-a7ec-9fe4e44f41a3%28Office.15%29.aspx)|
|[ForeignType](http://msdn.microsoft.com/library/a6cda280-bf0c-b8b0-0750-0ec5fbad90e0%28Office.15%29.aspx)|
|[FromConnects](http://msdn.microsoft.com/library/feb80221-c5d9-f72e-2f79-5153ed375383%28Office.15%29.aspx)|
|[GeometryCount](http://msdn.microsoft.com/library/4dffe649-3629-6e3e-bcc0-d860eb1efdbe%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/12784797-c42b-deee-9ae1-6115cd014ac8%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/c1f04a6f-032b-f626-c2e9-6688528052f6%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/948982c0-a872-802f-a2d3-69c6539ca3f2%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/7fb67e8b-76a7-c2ac-e7aa-89635ca7622c%28Office.15%29.aspx)|
|[IsCallout](http://msdn.microsoft.com/library/6977e383-41c5-effe-4ac9-88cfc0476813%28Office.15%29.aspx)|
|[IsDataGraphicCallout](http://msdn.microsoft.com/library/dedf6880-e597-8582-12e5-18bfe6286e66%28Office.15%29.aspx)|
|[IsOpenForTextEdit](http://msdn.microsoft.com/library/6a4525f2-2532-083d-87f7-390ae7034a65%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/6c7ab4ca-8813-9cbc-d433-a3991a0b450f%28Office.15%29.aspx)|
|[Layer](http://msdn.microsoft.com/library/fb076dda-fa1f-a1fe-c97b-03ba3c7041f0%28Office.15%29.aspx)|
|[LayerCount](http://msdn.microsoft.com/library/0ebcdf53-ebf3-8e26-236f-086f2c9f3c08%28Office.15%29.aspx)|
|[LengthIU](http://msdn.microsoft.com/library/11d57f17-5285-6b45-1da1-dc58db087395%28Office.15%29.aspx)|
|[LineStyle](http://msdn.microsoft.com/library/1d1f2b2e-705d-6547-f6d6-0c5693e426d6%28Office.15%29.aspx)|
|[LineStyleKeepFmt](http://msdn.microsoft.com/library/4dd4ee1e-5201-1602-39f1-bcda85f96bd0%28Office.15%29.aspx)|
|[Master](http://msdn.microsoft.com/library/698e205b-3cfc-2ee1-4fa1-73bc3d018b78%28Office.15%29.aspx)|
|[MasterShape](http://msdn.microsoft.com/library/bf710d8b-11f6-145d-a306-658dc23dedbf%28Office.15%29.aspx)|
|[MemberOfContainers](http://msdn.microsoft.com/library/e8ed57eb-4031-5718-07ce-3641bda00186%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/a0708af0-a813-7539-c43f-049009f1ab62%28Office.15%29.aspx)|
|[NameID](http://msdn.microsoft.com/library/ae658ed9-124f-22f2-53be-5c9b6ebaa382%28Office.15%29.aspx)|
|[NameU](http://msdn.microsoft.com/library/1f645016-86a5-f8e4-d5e4-00b8d02cc523%28Office.15%29.aspx)|
|[Object](http://msdn.microsoft.com/library/a2e8644a-ac7b-1bb7-9b6b-1515fb9126d2%28Office.15%29.aspx)|
|[ObjectIsInherited](http://msdn.microsoft.com/library/5bb91e7a-f28e-f169-2e4a-87d46aacdccc%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/d5711c8e-14a5-6e6b-e8f4-5501a483c9b9%28Office.15%29.aspx)|
|[OneD](http://msdn.microsoft.com/library/f1511393-4402-ecf8-82a2-2026c56622d0%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/aada0bc1-75e3-8357-3ef9-597a10250860%28Office.15%29.aspx)|
|[Paths](http://msdn.microsoft.com/library/8a179059-7cab-728a-c7b8-a4d8b31476ee%28Office.15%29.aspx)|
|[PathsLocal](http://msdn.microsoft.com/library/aa5da0de-ca06-69c0-1fbf-b19ea02d0088%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/6bfa4b18-b4f3-0ac0-de21-ed18600ff473%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/0ccd2df9-fd84-dee0-0d89-5b7115e418d6%28Office.15%29.aspx)|
|[ProgID](http://msdn.microsoft.com/library/2cd96dd5-7d73-77ea-9e7e-3d1dcd98a21a%28Office.15%29.aspx)|
|[RootShape](http://msdn.microsoft.com/library/c2e91d43-4968-cfee-e53b-4df115d171f6%28Office.15%29.aspx)|
|[RowCount](http://msdn.microsoft.com/library/358f07c8-f72a-134a-53d8-9b70f2400484%28Office.15%29.aspx)|
|[RowExists](http://msdn.microsoft.com/library/bd89deb9-eda3-18d8-6305-bd380d9e649f%28Office.15%29.aspx)|
|[RowsCellCount](http://msdn.microsoft.com/library/bb9c1990-5ead-e56b-7b09-a49a2b7ad111%28Office.15%29.aspx)|
|[RowType](http://msdn.microsoft.com/library/416b77f1-6cec-de5b-c2b8-c6e5b239c54c%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/e87823aa-fd7c-e222-417b-a167d2e0898a%28Office.15%29.aspx)|
|[SectionExists](http://msdn.microsoft.com/library/588a3b17-4831-b7bb-455f-12bc5c3620fc%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/83fea91a-19a6-f600-7d03-ba2f03f62d28%28Office.15%29.aspx)|
|[SpatialNeighbors](http://msdn.microsoft.com/library/98069519-d788-c34f-ac25-64bda73324d5%28Office.15%29.aspx)|
|[SpatialRelation](http://msdn.microsoft.com/library/7e9f26b5-2887-493f-01c1-5e3900ea8c05%28Office.15%29.aspx)|
|[SpatialSearch](http://msdn.microsoft.com/library/360b48b0-783a-7282-b3fe-83f424c393d4%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/c9d9d8bf-6e64-5231-b870-fcc5de7fdc7b%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/beba03ba-6926-d2db-4e36-652d05c2925c%28Office.15%29.aspx)|
|[StyleKeepFmt](http://msdn.microsoft.com/library/22403064-fa5d-c0cf-8ee7-0f8ae2895d3c%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/5c002c5d-f5ce-7f89-d799-4fc6ccb1a1f7%28Office.15%29.aspx)|
|[TextStyle](http://msdn.microsoft.com/library/9436ba1b-f792-aed6-3936-b2d88a6dd2ea%28Office.15%29.aspx)|
|[TextStyleKeepFmt](http://msdn.microsoft.com/library/add41319-8b81-a803-46e2-697df37eb731%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/0d7438d2-e2df-2045-1a2f-608eca530bc1%28Office.15%29.aspx)|
|[UniqueID](http://msdn.microsoft.com/library/a82e1175-4536-8919-6531-593d57c3b2f5%28Office.15%29.aspx)|

