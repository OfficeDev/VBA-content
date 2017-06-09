---
title: Selection Object (Visio)
keywords: vis_sdr.chm10220
f1_keywords:
- vis_sdr.chm10220
ms.prod: visio
api_name:
- Visio.Selection
ms.assetid: e5734140-6dbe-7de8-9695-1a22fb4ac628
ms.date: 06/08/2017
---


# Selection Object (Visio)

Represents a subset of  **Shape** objects for a page or master to which an operation can be applied.


## Remarks

To retrieve a  **Selection** object that corresponds to the set of shapes selected in a window, use the **Selection** property of a **Window** object.

The default property of a  **Selection** object is **Item**.

After you retrieve a  **Selection** object, you can add or remove shapes by using the **Select** method.

By default, the items reported by a  **Selection** object do not include subselected or superselected **Shape** objects. Use the **IterationMode** property to control whether subselected and superselected **Shape** objects are reported. You can determine whether an individual item is subselected or superselected by using the **ItemStatus** property.


## Methods



|**Name**|
|:-----|
|[AddToContainers](http://msdn.microsoft.com/library/7f3e739f-a573-049c-9f54-9e93a401191f%28Office.15%29.aspx)|
|[AddToGroup](http://msdn.microsoft.com/library/8bef7960-271c-245d-dec0-eeea4af66097%28Office.15%29.aspx)|
|[Align](http://msdn.microsoft.com/library/4a73dfee-2a78-f459-4481-5f722feb7204%28Office.15%29.aspx)|
|[AutomaticLink](http://msdn.microsoft.com/library/6943b2b1-269a-7759-d981-a3749cfbeaee%28Office.15%29.aspx)|
|[AvoidPageBreaks](http://msdn.microsoft.com/library/c0255ebe-5094-1196-0bfb-2693efefe47c%28Office.15%29.aspx)|
|[BoundingBox](http://msdn.microsoft.com/library/5ec076c3-5720-9215-16ef-8da0e674f86f%28Office.15%29.aspx)|
|[BreakLinkToData](http://msdn.microsoft.com/library/83a52ed7-1d10-9005-4a1a-339995106d8b%28Office.15%29.aspx)|
|[BringForward](http://msdn.microsoft.com/library/d12a81a5-6faa-6828-bdf0-279c27c89571%28Office.15%29.aspx)|
|[BringToFront](http://msdn.microsoft.com/library/f7e0b949-9f16-e4c1-8443-941abd3495db%28Office.15%29.aspx)|
|[Combine](http://msdn.microsoft.com/library/a74b25b0-6957-2088-f34f-4000c2be9736%28Office.15%29.aspx)|
|[ConnectShapes](http://msdn.microsoft.com/library/40e9c839-69f0-2142-6b9c-249212e373a4%28Office.15%29.aspx)|
|[ConvertToGroup](http://msdn.microsoft.com/library/bfd06685-bb44-b605-251f-334118fa11e7%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/e7d9ab14-7e64-f1fa-7813-62caee133b57%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/1f5d6f8a-81ab-3948-870c-a46a21f6b005%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/be259027-9cc4-95a4-2aa9-349b1967b9be%28Office.15%29.aspx)|
|[DeleteEx](http://msdn.microsoft.com/library/8935a2de-2fab-0b2e-1595-a78d3dc2fd90%28Office.15%29.aspx)|
|[DeselectAll](http://msdn.microsoft.com/library/2453beb9-e871-ef77-d420-2430c5466f8e%28Office.15%29.aspx)|
|[Distribute](http://msdn.microsoft.com/library/7750167b-b4ef-c1b6-68f4-1f40ab1fd33e%28Office.15%29.aspx)|
|[DrawRegion](http://msdn.microsoft.com/library/3c3a04d9-a275-a73e-8325-eadd3cae1999%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/515b522c-8b99-ea51-822f-47f0de24d330%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/41ecd499-358d-804a-3311-43d0041a5562%28Office.15%29.aspx)|
|[FitCurve](http://msdn.microsoft.com/library/d0f3c799-c15d-cdc8-c0b0-34aeeecec495%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/40ad506b-e5e2-4a42-6b38-0363e462fce4%28Office.15%29.aspx)|
|[FlipHorizontal](http://msdn.microsoft.com/library/97cecbcf-8489-c8b9-046e-28599f491e3c%28Office.15%29.aspx)|
|[FlipVertical](http://msdn.microsoft.com/library/e83d7faa-25c2-cdf2-ea78-de9061e5098a%28Office.15%29.aspx)|
|[Fragment](http://msdn.microsoft.com/library/e648675f-e60a-6a21-182e-32aa913df335%28Office.15%29.aspx)|
|[GetCallouts](http://msdn.microsoft.com/library/29adcbbc-d5a9-a284-c025-785ad1ccf2c8%28Office.15%29.aspx)|
|[GetContainers](http://msdn.microsoft.com/library/8e04bed5-f9ef-04bf-3013-c6dd623f9f63%28Office.15%29.aspx)|
|[GetIDs](http://msdn.microsoft.com/library/79b1fb3f-eb53-2640-a988-6e79b067f228%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/79afc3c4-7350-2196-7a07-3b7c5629568a%28Office.15%29.aspx)|
|[Intersect](http://msdn.microsoft.com/library/5dc63a77-62de-3892-6ed2-bcb5cb0a29f1%28Office.15%29.aspx)|
|[Join](http://msdn.microsoft.com/library/e176abcc-edd1-0e40-afc8-e05ed8dec998%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/58ff8c1f-92b3-2473-d786-28e64e7c5586%28Office.15%29.aspx)|
|[LayoutChangeDirection](http://msdn.microsoft.com/library/1c40348c-1884-1501-3609-aebf2e87686c%28Office.15%29.aspx)|
|[LayoutIncremental](http://msdn.microsoft.com/library/cae92d61-7800-a836-7e57-6d238661b02a%28Office.15%29.aspx)|
|[LinkToData](http://msdn.microsoft.com/library/1aa42548-2f3a-015d-e618-c0e103ffaea3%28Office.15%29.aspx)|
|[MemberOfContainersIntersection](http://msdn.microsoft.com/library/574282fa-3f1b-0e6a-a800-01ce447643f9%28Office.15%29.aspx)|
|[MemberOfContainersUnion](http://msdn.microsoft.com/library/b21b01df-08cd-4222-7ccd-1e2b9b34d462%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/12e60f50-f06d-45bb-b79d-db2e0d767461%28Office.15%29.aspx)|
|[MoveToSubprocess](http://msdn.microsoft.com/library/a61f1e93-06a3-6ddc-8cae-f92212078c96%28Office.15%29.aspx)|
|[Offset](http://msdn.microsoft.com/library/69eb7288-0540-18aa-9c71-96735018442e%28Office.15%29.aspx)|
|[RemoveFromContainers](http://msdn.microsoft.com/library/d1ed1360-3caa-3e03-98ef-84f4bd52a035%28Office.15%29.aspx)|
|[RemoveFromGroup](http://msdn.microsoft.com/library/4e593510-9970-c6fb-f598-e9f2e237bcb2%28Office.15%29.aspx)|
|[ReplaceShape](http://msdn.microsoft.com/library/dc278901-77ce-e1fe-c44f-f464bbb1c360%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/4fc41631-adb4-9c5a-570f-e8ccaa2701eb%28Office.15%29.aspx)|
|[ReverseEnds](http://msdn.microsoft.com/library/9175b098-6e1f-6b10-b685-d63896b397fc%28Office.15%29.aspx)|
|[Rotate](http://msdn.microsoft.com/library/3c0a1a4d-a172-131a-9fb4-d215a5b9b2af%28Office.15%29.aspx)|
|[Rotate90](http://msdn.microsoft.com/library/619f0b7f-027f-5cd6-361a-ec3db73a2712%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/b135632a-1158-1903-0b29-931c88deae21%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/e2280c51-84e8-4403-1c9e-f3bc504aff2f%28Office.15%29.aspx)|
|[SendBackward](http://msdn.microsoft.com/library/645a5686-6421-f8dd-425f-3cb5b0b7de85%28Office.15%29.aspx)|
|[SendToBack](http://msdn.microsoft.com/library/00417838-455b-c915-8879-64a83b0f1233%28Office.15%29.aspx)|
|[SetContainerFormat](http://msdn.microsoft.com/library/b0766138-07da-4539-b254-7692529e0771%28Office.15%29.aspx)|
|[SetQuickStyle](http://msdn.microsoft.com/library/39b810b5-0738-daed-0103-8a2df07559c6%28Office.15%29.aspx)|
|[Subtract](http://msdn.microsoft.com/library/606798b6-3482-0c45-d583-4762ee07da45%28Office.15%29.aspx)|
|[SwapEnds](http://msdn.microsoft.com/library/515580db-4018-30b3-0ed6-cb3a412b62c7%28Office.15%29.aspx)|
|[Trim](http://msdn.microsoft.com/library/0063d29a-3e47-bb2b-71fd-328c19a0a65b%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/b9f14342-e885-1399-83ed-59189f5cbec3%28Office.15%29.aspx)|
|[Union](http://msdn.microsoft.com/library/1ab7ce2a-98af-c455-7558-6f4f9226eeb9%28Office.15%29.aspx)|
|[UpdateAlignmentBox](http://msdn.microsoft.com/library/d7f13dcd-3ff6-0e0f-d996-afe59c16f813%28Office.15%29.aspx)|
|[VisualBoundingBox](http://msdn.microsoft.com/library/ae107bd8-ac99-6303-2820-a5afb19165a3%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/09aed34b-c509-33d7-efd5-7ac5d5b05482%28Office.15%29.aspx)|
|[ContainingMaster](http://msdn.microsoft.com/library/9eae609f-2d55-2180-ea9b-cf1f8ec7b7b3%28Office.15%29.aspx)|
|[ContainingMasterID](http://msdn.microsoft.com/library/9f9aad28-3e77-8ef8-29dc-e53852adf63d%28Office.15%29.aspx)|
|[ContainingPage](http://msdn.microsoft.com/library/dca54861-d6c6-9d39-2a49-2070a578607f%28Office.15%29.aspx)|
|[ContainingPageID](http://msdn.microsoft.com/library/f7d19685-9e1d-8867-978a-563dd3e93b0b%28Office.15%29.aspx)|
|[ContainingShape](http://msdn.microsoft.com/library/c25dec03-dfa9-d61f-ad02-8ea7ee6cd87f%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/89432479-5457-838f-a85d-20eb0dd61547%28Office.15%29.aspx)|
|[DataGraphic](http://msdn.microsoft.com/library/09275500-7b8a-2d78-971c-2e27bc3b9e46%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/fa7d64c9-1d50-3e35-cece-32b52790d158%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/dee5994c-d43b-1833-1ea0-17fc24f01d74%28Office.15%29.aspx)|
|[FillStyle](http://msdn.microsoft.com/library/efdf51ba-7d0a-d5c0-5a39-d22d7a79a053%28Office.15%29.aspx)|
|[FillStyleKeepFmt](http://msdn.microsoft.com/library/e4034e7d-3a81-3fe6-0fb5-61549942c8cb%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/3f09566d-eec6-0c20-87bc-60db45d3e23f%28Office.15%29.aspx)|
|[ItemStatus](http://msdn.microsoft.com/library/2dcd9875-222d-fdb9-c2be-1a1df4ee86e7%28Office.15%29.aspx)|
|[IterationMode](http://msdn.microsoft.com/library/e4cd372c-a156-364d-f051-d9a8c618bd2c%28Office.15%29.aspx)|
|[LineStyle](http://msdn.microsoft.com/library/8bfba446-5987-58d1-54e2-5e861d7ce48d%28Office.15%29.aspx)|
|[LineStyleKeepFmt](http://msdn.microsoft.com/library/63703d4e-34b6-9b53-c2c1-b7503d0c3986%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/b21e23b1-8ff3-ec9e-f92d-230f0ea250a7%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/a9e513e8-386a-99c8-6d7e-b525c6dc8b54%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/fb3e59d3-2739-beee-441c-ffcee6621aa0%28Office.15%29.aspx)|
|[PrimaryItem](http://msdn.microsoft.com/library/febdc4ec-d7db-7b4f-145b-aa9b23a2d5d2%28Office.15%29.aspx)|
|[SelectionForDragCopy](http://msdn.microsoft.com/library/f7e6e87a-c904-6008-fdde-4d5cb124351c%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/cd7ecc8b-8513-d901-9f86-670569e53a4b%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/f0853c43-14b4-bcd9-eb07-fbc0312e106b%28Office.15%29.aspx)|
|[StyleKeepFmt](http://msdn.microsoft.com/library/b56bfda8-0076-0114-b231-bb7c649c6310%28Office.15%29.aspx)|
|[TextStyle](http://msdn.microsoft.com/library/3b94d8a1-e3aa-0473-de85-744cb353886e%28Office.15%29.aspx)|
|[TextStyleKeepFmt](http://msdn.microsoft.com/library/d9900f73-dc39-e717-d923-78a9b275271e%28Office.15%29.aspx)|

