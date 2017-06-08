---
title: ShapeRange Object (PowerPoint)
keywords: vbapp10.chm548000
f1_keywords:
- vbapp10.chm548000
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange
ms.assetid: 0a194183-380e-ffb6-9336-b5bd311e917d
ms.date: 06/08/2017
---


# ShapeRange Object (PowerPoint)

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as a single shape or as many as all the shapes on the document.


## Remarks

You can include whichever shapes you want — chosen from among all the shapes on the document or all the shapes in the selection — to construct a shape range. For example, you could construct a  **ShapeRange** collection that contains the first three shapes on a document, all the selected shapes on a document, or all the freeforms on a document.

For an overview of how to work with either a single shape or with more than one shape at a time, see [How to: Work with Shapes (Drawing Objects)](http://msdn.microsoft.com/library/3ffaaaea-6406-262b-2bc7-788699175266%28Office.15%29.aspx).

The following examples describe how to:


- Return a set of shapes you specify by name or index number.
    
- Return all or some of the selected shapes on a document.
    

## Example

Use  **Shapes.Range** (index), where index is the name or index number of the shape or an array that contains either names or index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes on a document. You can use the **Array** function to construct an array of names or index numbers. The following example sets the fill pattern for shapes one and three on `myDocument`.


```
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Range(Array(1, 3)).Fill _

    .Patterned msoPatternHorizontalBrick
```

The following example sets the fill pattern for the shapes named "Oval 4" and "Rectangle 5" on  `myDocument`.




```
Set myDocument = ActivePresentation.Slides(1)

Set myRange = myDocument.Shapes _

    .Range(Array("Oval 4", "Rectangle 5"))

myRange.Fill.Patterned msoPatternHorizontalBrick
```

Although you can use the [Range](http://msdn.microsoft.com/library/5ee926d9-5b30-a26b-7365-f4709a1a7bdb%28Office.15%29.aspx)method to return any number of shapes or slides, it is simpler to use the [Item](http://msdn.microsoft.com/library/c93d269b-7405-6e3f-6d59-d1e18e7f0cb1%28Office.15%29.aspx)method if you want to return only a single member of the collection. For example,  `Shapes(1)` is simpler than `Shapes.Range(1)`.

Use the [ShapeRange](http://msdn.microsoft.com/library/3fd7aed0-ab63-adaa-1a46-c745b6c3e245%28Office.15%29.aspx)property of the  **Selection** object to return all the shapes in the selection. The following example sets the fill foreground color for all the shapes in the selection in window one, assuming that there's at least one shape in the selection.




```
Windows(1).Selection.ShapeRange.Fill.ForeColor _

    .RGB = RGB(255, 0, 255)
```

Use  **Selection.ShapeRange** (index), where index is the shape name or the index number, to return a single shape within the selection. The following example sets the fill foreground color for shape two in the collection of selected shapes in window one, assuming that there are at least two shapes in the selection.




```
Windows(1).Selection.ShapeRange(2).Fill.ForeColor _

    .RGB = RGB(255, 0, 255)
```


## Methods



|**Name**|
|:-----|
|[Align](http://msdn.microsoft.com/library/5d4553ad-521a-1f3c-77ba-3dd5fbd02a09%28Office.15%29.aspx)|
|[Apply](http://msdn.microsoft.com/library/4aabc231-0854-070e-01d3-5ac48d16afbd%28Office.15%29.aspx)|
|[ApplyAnimation](http://msdn.microsoft.com/library/cfaa7d9c-3a65-1be7-dd6c-61e01b9e7d36%28Office.15%29.aspx)|
|[ConvertTextToSmartArt](http://msdn.microsoft.com/library/c61b8cb6-d5a6-00bf-49e6-94b7a9499e75%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/ddc0dad9-6647-e2f4-393a-347c273656dd%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/0e86d67c-7d52-4f3a-4cdd-6363667600a1%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/9c3245de-828c-8a54-d267-74d41a841509%28Office.15%29.aspx)|
|[Distribute](http://msdn.microsoft.com/library/bbabe9db-30ba-e165-0dcc-7a15e849228e%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/da7e1e45-480d-313d-1d12-65ee5bf26d86%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/e9f5ceb5-2ddf-d70c-41d5-d5877043b62a%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/f70f3986-3a39-78f9-476e-b72ef000c469%28Office.15%29.aspx)|
|[IncrementLeft](http://msdn.microsoft.com/library/08d84101-bdfe-c3c6-a309-00c2fb2adab5%28Office.15%29.aspx)|
|[IncrementRotation](http://msdn.microsoft.com/library/427367bb-5264-86de-cf39-be252c4b7098%28Office.15%29.aspx)|
|[IncrementTop](http://msdn.microsoft.com/library/55c18051-97a8-beab-c354-48256daff762%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/c93d269b-7405-6e3f-6d59-d1e18e7f0cb1%28Office.15%29.aspx)|
|[MergeShapes](http://msdn.microsoft.com/library/fea16a4d-9ee2-83fb-e5f5-00640d133d3b%28Office.15%29.aspx)|
|[PickUp](http://msdn.microsoft.com/library/f583c44c-0ab1-19eb-40f7-7e3412c93686%28Office.15%29.aspx)|
|[PickupAnimation](http://msdn.microsoft.com/library/13210009-1329-8c3e-01ce-459e1bcac88c%28Office.15%29.aspx)|
|[Regroup](http://msdn.microsoft.com/library/3da4a44d-4b0c-e335-b376-4d76fe5ed561%28Office.15%29.aspx)|
|[RerouteConnections](http://msdn.microsoft.com/library/61db5f5d-74cd-1b9d-1b37-9d33e320cca8%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/3e86cfd8-1df6-a164-d19b-8d53b7b52dc0%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/868f56cb-6a3a-902e-b6a9-2a9229936b41%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/475f035e-a266-c263-eb62-542c51bb4087%28Office.15%29.aspx)|
|[SetShapesDefaultProperties](http://msdn.microsoft.com/library/169f174a-1e2a-370e-663c-08a851f1e4d3%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/7bac0e8b-09d5-b219-af20-2a3b8dcee9d9%28Office.15%29.aspx)|
|[UpgradeMedia](http://msdn.microsoft.com/library/a05e171a-1fff-1128-7a2d-a5576593fc70%28Office.15%29.aspx)|
|[ZOrder](http://msdn.microsoft.com/library/906620bd-9293-694a-002d-97e760de988a%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActionSettings](http://msdn.microsoft.com/library/5e4c3e26-be69-ce78-41e4-903534fde7a9%28Office.15%29.aspx)|
|[Adjustments](http://msdn.microsoft.com/library/e76f2051-c362-9848-59be-fc3c9662e3a8%28Office.15%29.aspx)|
|[AlternativeText](http://msdn.microsoft.com/library/5255de02-810d-981b-4b2d-9a28fbcdae4c%28Office.15%29.aspx)|
|[AnimationSettings](http://msdn.microsoft.com/library/b248113c-54f6-5a36-b77f-63d76c10f7f3%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/6d226806-1452-3a6b-6a0f-ccf0ea0626c7%28Office.15%29.aspx)|
|[AutoShapeType](http://msdn.microsoft.com/library/a1b6c923-dac7-8b5a-6d8b-46a62cfb119e%28Office.15%29.aspx)|
|[BackgroundStyle](http://msdn.microsoft.com/library/5abc16be-d2b1-0138-49be-227dd467869f%28Office.15%29.aspx)|
|[BlackWhiteMode](http://msdn.microsoft.com/library/a9d51d2d-aee3-78ba-3213-6ad7263f268c%28Office.15%29.aspx)|
|[Callout](http://msdn.microsoft.com/library/ccff61a0-d077-a80d-d1ce-be9b036842c0%28Office.15%29.aspx)|
|[Chart](http://msdn.microsoft.com/library/15b69ed5-db0e-0bae-403d-263eedb7b4a1%28Office.15%29.aspx)|
|[Child](http://msdn.microsoft.com/library/882bb89f-1a0c-4384-55cd-4399f4e840c0%28Office.15%29.aspx)|
|[ConnectionSiteCount](http://msdn.microsoft.com/library/352f9c7c-6290-f974-5924-01e108fb4919%28Office.15%29.aspx)|
|[Connector](http://msdn.microsoft.com/library/04871183-d9d0-f0ba-f801-4f1f6564f0d3%28Office.15%29.aspx)|
|[ConnectorFormat](http://msdn.microsoft.com/library/30d41f5e-3bd5-b8ed-cba9-9de8b7567a98%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/17d38ae2-667c-d256-2098-4ed110b7488f%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/6c273206-ecd1-d420-bf40-877ca678876c%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/b1515a2f-e701-17ec-9224-77af548b002f%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/689cef96-6ad8-aa20-27c6-065af06b5753%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/d2080e84-8876-ab45-330d-718ed1bc505f%28Office.15%29.aspx)|
|[GroupItems](http://msdn.microsoft.com/library/94d0e684-5237-2415-e222-cd38cbd22e36%28Office.15%29.aspx)|
|[HasChart](http://msdn.microsoft.com/library/b863fc82-6f99-d102-a399-fde44af9e877%28Office.15%29.aspx)|
|[HasInkXML](http://msdn.microsoft.com/library/1a7b7b8b-c0e8-8f62-1015-e99cb590fd50%28Office.15%29.aspx)|
|[HasSmartArt](http://msdn.microsoft.com/library/9c207244-c829-549a-aebc-aa768ac12ecd%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/aaf47e4f-0315-2311-e9c5-68a12d36235c%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/6359fbeb-0a91-ad56-9edd-9b6be7fe51b7%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/d70eeb9d-d3d2-51ee-1567-f8762aaa089b%28Office.15%29.aspx)|
|[HorizontalFlip](http://msdn.microsoft.com/library/4c41e250-2a8f-3eab-3244-0910fb43362e%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/9bc2df4a-441f-27fa-c808-1e87b2a4be7e%28Office.15%29.aspx)|
|[InkXML](http://msdn.microsoft.com/library/faff227c-293a-58cf-fe49-eb8b5f5caac3%28Office.15%29.aspx)|
|[IsNarration](http://msdn.microsoft.com/library/a82b4156-9025-aa7c-b132-b7f5cafa2b3b%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/eb27c0ea-68d1-4819-5708-fa5a198cc086%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/27f648e0-d7eb-27a9-312b-8aa1784e7001%28Office.15%29.aspx)|
|[LinkFormat](http://msdn.microsoft.com/library/aa2f91d3-b3fd-9834-b189-ec7fc783db6d%28Office.15%29.aspx)|
|[LockAspectRatio](http://msdn.microsoft.com/library/e30f2834-b6c2-d966-dbee-b22912e4e3f0%28Office.15%29.aspx)|
|[MediaFormat](http://msdn.microsoft.com/library/d8c02203-9570-247c-d0c4-d823b349ad84%28Office.15%29.aspx)|
|[MediaType](http://msdn.microsoft.com/library/4d3d321c-6af5-36ce-5bf8-363dfce1a05f%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/b87c7def-f68d-0dde-e971-2201043f6bfc%28Office.15%29.aspx)|
|[Nodes](http://msdn.microsoft.com/library/da1e20b4-4c03-9d7c-8f01-9373ad97ca77%28Office.15%29.aspx)|
|[OLEFormat](http://msdn.microsoft.com/library/ff454e81-5c55-5deb-9816-0eb06b236a95%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/d43d43e8-8b92-bf87-fc4e-160166f26b10%28Office.15%29.aspx)|
|[ParentGroup](http://msdn.microsoft.com/library/425aec51-78d8-8e44-7d33-a300af184676%28Office.15%29.aspx)|
|[PictureFormat](http://msdn.microsoft.com/library/5d51631d-1cd4-fbfc-9198-9d883b281821%28Office.15%29.aspx)|
|[PlaceholderFormat](http://msdn.microsoft.com/library/3c3c344f-aa02-29b2-5ef5-d090f3e32a2c%28Office.15%29.aspx)|
|[Reflection](http://msdn.microsoft.com/library/1ed3bdf2-e02f-994c-5556-c7b33658a9c6%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/06969cb4-086d-360e-70eb-5e7a80da5f69%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/01aa0a5a-341b-6764-e3ea-1f20379d0de3%28Office.15%29.aspx)|
|[ShapeStyle](http://msdn.microsoft.com/library/7809d2e9-091f-acde-0eaa-130031e5d5bc%28Office.15%29.aspx)|
|[SmartArt](http://msdn.microsoft.com/library/e91922e1-71a6-009e-4592-cbd7f5934270%28Office.15%29.aspx)|
|[SoftEdge](http://msdn.microsoft.com/library/5a61651f-0935-cca0-e5dd-c0b71fb703c4%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/2ab10bd4-071a-8e84-cf46-1687e6661bb8%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/98620c36-50aa-a2a0-e6b6-125229dd87af%28Office.15%29.aspx)|
|[TextEffect](http://msdn.microsoft.com/library/8cf70ead-8534-ef82-5064-21b9929e6f08%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/ec6093f2-232b-361b-b85d-7a99fafd8878%28Office.15%29.aspx)|
|[TextFrame2](http://msdn.microsoft.com/library/56554e58-c16b-09dd-8acc-4e2da7715bef%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/e0e2f72d-639b-86fd-2191-f537ddcd45ad%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/bb4e08a3-6517-c500-23ac-ec65b3340f76%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/448b4c64-6519-ce0d-fb2e-9dbc65462494%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/bad9a4a8-267a-cfb5-e990-66bf751e5814%28Office.15%29.aspx)|
|[VerticalFlip](http://msdn.microsoft.com/library/868657a8-72c6-896d-6a6f-f9adbbc88a59%28Office.15%29.aspx)|
|[Vertices](http://msdn.microsoft.com/library/3ded99dc-f64d-cfdd-f982-2e892ba4a446%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/39ca5142-bf7b-a48d-ce9d-e929e4611aac%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/b9b521f8-70e0-90aa-fdbf-675c78cc0d28%28Office.15%29.aspx)|
|[ZOrderPosition](http://msdn.microsoft.com/library/606b0140-086d-54ec-fdbf-16edf38e5170%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
