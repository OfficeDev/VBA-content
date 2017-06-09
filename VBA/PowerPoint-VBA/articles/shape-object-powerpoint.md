---
title: Shape Object (PowerPoint)
keywords: vbapp10.chm547000
f1_keywords:
- vbapp10.chm547000
ms.prod: powerpoint
api_name:
- PowerPoint.Shape
ms.assetid: 1da93849-99e0-827e-ced3-c6cf7f8569f3
ms.date: 06/08/2017
---


# Shape Object (PowerPoint)

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


## Remarks


 **Note**  There are three objects that represent shapes: the  **Shapes** collection, which represents all the shapes on a document; the **[ShapeRange](http://msdn.microsoft.com/library/0a194183-380e-ffb6-9336-b5bd311e917d%28Office.15%29.aspx)** collection, which represents a specified subset of the shapes on a document (for example, a **ShapeRange** object could represent shapes one and four on the document, or it could represent all the selected shapes on the document); the **Shape** object, which represents a single shape on a document. If you want to work with several shape at the same time or with shapes within the selection, use a **ShapeRange** collection. For an overview of how to work with either a single shape or with more than one shape at a time, see [How to: Work with Shapes (Drawing Objects)](http://msdn.microsoft.com/library/3ffaaaea-6406-262b-2bc7-788699175266%28Office.15%29.aspx).

The following examples describe how to:


- Return an existing shape on a slide, indexed by name or number.
    
- Return a newly created shape on a slide.
    
- Return a shape within the selection.
    
- Return the slide title and other placeholders on a slide.
    
- Return the shapes attached to the ends of a connector.
    
- Return the default shape for a presentation.
    
- Return a newly created freeform.
    
- Return a single shape from within a group.
    
- Return a newly formed group of shapes.
    

## Example

Use  **Shapes** (index), where index is the shape name or the index number, to return a **Shape** object that represents a shape on a slide. The following example horizontally flips shape one and the shape named Rectangle 1 on myDocument.


```
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).Flip msoFlipHorizontal

myDocument.Shapes("Rectangle 1").Flip msoFlipHorizontal
```

Each shape is assigned a default name when you add it to the  **Shapes** collection. To give the shape a more meaningful name, use the **Name** property. The following example adds a rectangle to myDocument, gives it the name Red Square, and then sets its foreground color and line style.




```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(Type:=msoShapeRectangle, _

        Top:=144, Left:=144, Width:=72, Height:=72)

    .Name = "Red Square"

    .Fill.ForeColor.RGB = RGB(255, 0, 0)

    .Line.DashStyle = msoLineDashDot

End With
```

To add a shape to a slide and return a  **Shape** object that represents the newly created shape, use one of the following methods of the **Shapes** collection:[AddCallout](http://msdn.microsoft.com/library/e4b468d7-793a-09ae-fcfc-6a73db93c90e%28Office.15%29.aspx), [AddComment](http://msdn.microsoft.com/library/11347ca1-cef3-0923-2544-cb80e7fc5768%28Office.15%29.aspx), [AddConnector](http://msdn.microsoft.com/library/407eee86-11c1-7bee-ed25-aba71a930a1c%28Office.15%29.aspx), [AddCurve](http://msdn.microsoft.com/library/47f90182-a71b-a028-c43f-a85d59d2a56b%28Office.15%29.aspx), [AddLabel](http://msdn.microsoft.com/library/b744daf1-5b99-9649-8b97-d3f2193373c1%28Office.15%29.aspx), [AddLine](http://msdn.microsoft.com/library/9dbe640b-5ba4-a620-d3c6-4a2d0cc2bc27%28Office.15%29.aspx), [AddMediaObject](http://msdn.microsoft.com/library/7e2ab704-7fd4-86d7-3f61-8d69c13b5685%28Office.15%29.aspx), [AddOLEObject](http://msdn.microsoft.com/library/88a5aa63-0531-b9d8-43d2-5a995b91b189%28Office.15%29.aspx), [AddPicture](http://msdn.microsoft.com/library/af432432-b09b-3ca6-d392-132bd78251c7%28Office.15%29.aspx), [AddPlaceholder](http://msdn.microsoft.com/library/10927d59-1810-2f91-eb52-c42113570ccc%28Office.15%29.aspx), [AddPolyline](http://msdn.microsoft.com/library/e42c4f7a-de68-88bf-d250-28e642b56232%28Office.15%29.aspx), [AddShape](http://msdn.microsoft.com/library/2bc6cce5-3461-61ff-083d-bd36ee71cb59%28Office.15%29.aspx), [AddTable](http://msdn.microsoft.com/library/77ce193e-10f7-25f4-a6f8-99d7d2b781ad%28Office.15%29.aspx), [AddTextbox](http://msdn.microsoft.com/library/0c7c6093-48f6-e1f1-1837-e69d6ef13e57%28Office.15%29.aspx), [AddTextEffect](http://msdn.microsoft.com/library/4428ac57-c704-475a-1640-78a556e9ac3d%28Office.15%29.aspx), [AddTitle](http://msdn.microsoft.com/library/1fe13529-526a-1b29-7589-c155f9e46379%28Office.15%29.aspx).

Use  **Selection.ShapeRange** (index), where index is the shape name or the index number, to return a **Shape** object that represents a shape within the selection. The following example sets the fill for the first shape in the selection in the active window, assuming that there's at least one shape in the selection.




```
ActiveWindow.Selection.ShapeRange(1).Fill _

    .ForeColor.RGB = RGB(255, 0, 0)
```

Use  **Shapes.Title** to return a **Shape** object that represents an existing slide title. Use **Shapes.AddTitle** to add a title to a slide that doesn't already have one and return a **Shape** object that represents the newly created title. Use **Shapes.Placeholders** (index), where index is the placeholder's index number, to return a **Shape** object that represents a placeholder. If you have not changed the layering order of the shapes on a slide, the following three statements are equivalent, assuming that slide one has a title.




```
ActivePresentation.Slides(1).Shapes.Title _

    .TextFrame.TextRange.Font.Italic = True

ActivePresentation.Slides(1).Shapes.Placeholders(1) _

    .TextFrame.TextRange.Font.Italic = True

ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange.Font.Italic = True
```

To return a  **Shape** object that represents one of the shapes attached by a connector, use the[BeginConnectedShape](http://msdn.microsoft.com/library/7456899e-3f1c-3af8-e942-a6de1abeeca3%28Office.15%29.aspx)or [EndConnectedShape](http://msdn.microsoft.com/library/1d18fd9a-fc43-b50e-5c1a-dc6b5332b37e%28Office.15%29.aspx)property.



To return a  **Shape** object that represents the default shape for a presentation, use the[DefaultShape](http://msdn.microsoft.com/library/318ec04a-8b30-29b3-c8a6-732564efd7a8%28Office.15%29.aspx)property.



Use the [BuildFreeform](http://msdn.microsoft.com/library/330ea348-9f8c-c418-d67f-e4fd6c105c59%28Office.15%29.aspx)and [AddNodes](http://msdn.microsoft.com/library/4022d4cd-796b-8917-7265-d97bff5282ef%28Office.15%29.aspx)methods to define the geometry of a new freeform, and use the [ConvertToShape](http://msdn.microsoft.com/library/bc3d209e-6735-3011-9334-46049d269355%28Office.15%29.aspx)method to create the freeform and return the  **Shape** object that represents it.

Use  **GroupItems** (index), where index is the shape name or the index number within the group, to return a **Shape** object that represents a single shape in a grouped shape.

Use the [Group](http://msdn.microsoft.com/library/f70f3986-3a39-78f9-476e-b72ef000c469%28Office.15%29.aspx)or [Regroup](http://msdn.microsoft.com/library/3da4a44d-4b0c-e335-b376-4d76fe5ed561%28Office.15%29.aspx)method to group a range of shapes and return a single  **Shape** object that represents the newly formed group. After a group has been formed, you can work with the group the same way you work with any other shape.


## Methods



|**Name**|
|:-----|
|[Apply](http://msdn.microsoft.com/library/699a945f-656a-170a-e784-07b3004a858f%28Office.15%29.aspx)|
|[ApplyAnimation](http://msdn.microsoft.com/library/e3c65ffb-ea84-d5fd-4b14-25f517fb02f4%28Office.15%29.aspx)|
|[ConvertTextToSmartArt](http://msdn.microsoft.com/library/8ac35770-5835-c698-c0f1-12c3c03902c6%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/41c82fd1-9ee7-c937-0a75-77b84c33c972%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/908c998d-a15f-5075-33e0-de6c124a0ec7%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/998a345f-31e3-1270-7826-17d84d60634b%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/0d2f22bc-ee72-6405-011a-77a9eb98fb39%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/f340183a-4ef6-1a17-bbbb-5b1ec2b9aa4e%28Office.15%29.aspx)|
|[IncrementLeft](http://msdn.microsoft.com/library/a2b9a5e8-ba8c-612d-817f-c05d3df800b9%28Office.15%29.aspx)|
|[IncrementRotation](http://msdn.microsoft.com/library/f6e494fa-6bc1-0fc1-3bd3-ecc4fa5852e0%28Office.15%29.aspx)|
|[IncrementTop](http://msdn.microsoft.com/library/b74307f9-9efa-4117-9232-24dfd2bdb883%28Office.15%29.aspx)|
|[PickUp](http://msdn.microsoft.com/library/35730a7e-3878-dfae-2aba-3395d41e5f3e%28Office.15%29.aspx)|
|[PickupAnimation](http://msdn.microsoft.com/library/21068cec-c9c0-4a50-f318-74a9ff654091%28Office.15%29.aspx)|
|[RerouteConnections](http://msdn.microsoft.com/library/0928190d-d184-7522-1ce2-0fa884950220%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/0324449a-535c-e5ec-a9c3-0913f66057c0%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/2fc35ce6-62f5-7fa5-582d-26df91656a50%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/9fcf0ba4-ee6e-ecca-7948-7542db03ee99%28Office.15%29.aspx)|
|[SetShapesDefaultProperties](http://msdn.microsoft.com/library/4974cc1b-28af-94da-0821-76ffb698e2c3%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/2d0447df-7356-35e7-972e-e763ac1b3b8e%28Office.15%29.aspx)|
|[UpgradeMedia](http://msdn.microsoft.com/library/459ee25b-711f-2b74-38a0-3e209df7641b%28Office.15%29.aspx)|
|[ZOrder](http://msdn.microsoft.com/library/3317b5c3-611f-7cf8-3ce3-6d09255aa75f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActionSettings](http://msdn.microsoft.com/library/67e76de6-c0c3-7a35-f01e-e1cab4eb13d3%28Office.15%29.aspx)|
|[Adjustments](http://msdn.microsoft.com/library/2bb29847-cbeb-891b-c1e2-18e8ea7e8122%28Office.15%29.aspx)|
|[AlternativeText](http://msdn.microsoft.com/library/0ffde7b0-8a91-5456-e092-379491327a15%28Office.15%29.aspx)|
|[AnimationSettings](http://msdn.microsoft.com/library/c960d0de-afb3-55f2-b6fb-e67779cc42d2%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/e4e8fb64-0bb0-90c4-579c-f19c45030dfc%28Office.15%29.aspx)|
|[AutoShapeType](http://msdn.microsoft.com/library/99c8e48a-2e0e-0c5a-fb78-91790c668bd7%28Office.15%29.aspx)|
|[BackgroundStyle](http://msdn.microsoft.com/library/8b25d075-1ba8-ca90-7ec3-d28d7e7fa838%28Office.15%29.aspx)|
|[BlackWhiteMode](http://msdn.microsoft.com/library/bed5df5a-87b5-5e61-6d28-48a7776d0d83%28Office.15%29.aspx)|
|[Callout](http://msdn.microsoft.com/library/381f8eaa-f373-b1aa-6a53-4086d7e887d8%28Office.15%29.aspx)|
|[Chart](http://msdn.microsoft.com/library/7b641a32-a3e8-4d4f-3213-1e38ddb0efae%28Office.15%29.aspx)|
|[Child](http://msdn.microsoft.com/library/53371144-eabb-3f1f-f9cf-9a4e7b701d5f%28Office.15%29.aspx)|
|[ConnectionSiteCount](http://msdn.microsoft.com/library/2180bb96-d205-03f3-1ace-355f34286b2e%28Office.15%29.aspx)|
|[Connector](http://msdn.microsoft.com/library/3e8cc3be-30a6-4e4e-32ca-bfd55ae973c2%28Office.15%29.aspx)|
|[ConnectorFormat](http://msdn.microsoft.com/library/6c3f7f40-02a8-73ff-5829-7994ba1495d2%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/77d85e2f-aeba-7aba-b3d4-efe37ee487fe%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/ea2d0391-c093-09ec-ef45-01f0cd59db77%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/bfb2dfe6-5036-0731-3a0f-1294ba87e103%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/58bea564-b90a-4b39-53c7-8bb6f6ccd019%28Office.15%29.aspx)|
|[GroupItems](http://msdn.microsoft.com/library/295499de-0e74-e4ad-1145-f21927cbf2a9%28Office.15%29.aspx)|
|[HasChart](http://msdn.microsoft.com/library/5de934a4-03c2-959f-b0b9-562217146640%28Office.15%29.aspx)|
|[HasInkXML](http://msdn.microsoft.com/library/3d985f9b-64e3-8712-fd5f-73d38ca56810%28Office.15%29.aspx)|
|[HasSmartArt](http://msdn.microsoft.com/library/949d84a0-cdce-4351-70b7-f7dc92f3d5aa%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/fa38891a-e915-3a5c-4169-3c14e5e7136e%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/ea1a53e4-32d8-e51f-9e60-9ef719c0d973%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/3e2e7adf-9115-a903-c119-6429a10cbd9e%28Office.15%29.aspx)|
|[HorizontalFlip](http://msdn.microsoft.com/library/bf061a08-978c-dfb3-8a8f-4ecd62d95c53%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/b8d1c2ed-08e6-2a1d-7603-d80387fa4ee4%28Office.15%29.aspx)|
|[InkXML](http://msdn.microsoft.com/library/01e01d61-89a3-1314-fda5-6354d6590aa5%28Office.15%29.aspx)|
|[IsNarration](http://msdn.microsoft.com/library/e07e42e3-149d-153f-6852-a41c0eae80e3%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/1dffff64-fe2a-c164-52e2-2ea3507c6bec%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/edb5f40e-8b1e-fd3f-33da-0d4f1d465525%28Office.15%29.aspx)|
|[LinkFormat](http://msdn.microsoft.com/library/b742d78a-2fd3-1eb9-76d1-f2a2263cc68a%28Office.15%29.aspx)|
|[LockAspectRatio](http://msdn.microsoft.com/library/b66acf40-1136-36b6-eabc-96b0fac527de%28Office.15%29.aspx)|
|[MediaFormat](http://msdn.microsoft.com/library/e44c15c6-bfe4-010b-ab40-adc80e556da6%28Office.15%29.aspx)|
|[MediaType](http://msdn.microsoft.com/library/c42e3490-a4c9-d0bf-a201-71deff78d4b2%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/66e1d7e8-9398-8f01-d130-7206a48a63b3%28Office.15%29.aspx)|
|[Nodes](http://msdn.microsoft.com/library/85021d71-78f8-43e5-5a15-a0c1ae29ef61%28Office.15%29.aspx)|
|[OLEFormat](http://msdn.microsoft.com/library/d9353732-0b91-ae53-a468-07a57359295d%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/a88b1ec0-79de-4aef-9b71-a21bf8de2f44%28Office.15%29.aspx)|
|[ParentGroup](http://msdn.microsoft.com/library/1566110f-81dc-b73a-d658-2f6189113068%28Office.15%29.aspx)|
|[PictureFormat](http://msdn.microsoft.com/library/97d6b8d0-ddfb-c3b8-70fe-7569f5738f92%28Office.15%29.aspx)|
|[PlaceholderFormat](http://msdn.microsoft.com/library/4ccd4f93-74fc-be23-5ef4-0089d7247724%28Office.15%29.aspx)|
|[Reflection](http://msdn.microsoft.com/library/6120a828-e937-9b91-57c8-c896a4e2c9e9%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/3ed090a8-d945-85ee-155b-08628feff348%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/832b8e62-4fc5-1f4b-74c7-cc0e63a12699%28Office.15%29.aspx)|
|[ShapeStyle](http://msdn.microsoft.com/library/b93ffebd-8ace-6876-8336-96febb46be8c%28Office.15%29.aspx)|
|[SmartArt](http://msdn.microsoft.com/library/ac652436-8cdf-12a8-93c6-e50479e961b8%28Office.15%29.aspx)|
|[SoftEdge](http://msdn.microsoft.com/library/ec0e0555-a6fe-e389-e6b7-7ffa551e885b%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/cc57c50b-8c88-d863-31d2-a758eff5359f%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/23104a05-c2f0-21ea-f7ef-3bdc5587ce18%28Office.15%29.aspx)|
|[TextEffect](http://msdn.microsoft.com/library/b5d0a0a5-462d-1ede-3dac-7bedaaa1e318%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/6e4ad91e-c356-6a73-883d-8a0fd18c6ff6%28Office.15%29.aspx)|
|[TextFrame2](http://msdn.microsoft.com/library/bc76d1e5-3feb-51c9-a4d4-61f0bf36183f%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/16f0bc6a-ae6c-f4c3-9e3c-641f069eb7f6%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/fc675bc2-0af9-3f72-9b37-fabd586bbb2d%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/cf56f128-43d7-4f6e-f34c-83fbae854c12%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/3a6aa03d-8d93-9a08-ef42-8f128ada7b87%28Office.15%29.aspx)|
|[VerticalFlip](http://msdn.microsoft.com/library/56bf36e4-49df-5ae5-855c-3275d634dee4%28Office.15%29.aspx)|
|[Vertices](http://msdn.microsoft.com/library/b9ce441c-b305-4224-3fe8-3f947199a4d4%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/bf4d6dc9-fcae-1ae8-000f-736efcad34fc%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/b95213f9-2689-5131-5b85-d2eb661502a6%28Office.15%29.aspx)|
|[ZOrderPosition](http://msdn.microsoft.com/library/07e6c756-ae7b-f6d9-d903-92aef3b7fa9e%28Office.15%29.aspx)|

## See also


#### Concepts


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)

