---
title: ShapeRange Object (Publisher)
keywords: vbapb10.chm2359295
f1_keywords:
- vbapb10.chm2359295
ms.prod: publisher
api_name:
- Publisher.ShapeRange
ms.assetid: c85967c9-af43-747d-7e0b-64ddc22c84be
ms.date: 06/08/2017
---


# ShapeRange Object (Publisher)

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as one shape or as many as all the shapes in the document. You can include whichever shapes you want — chosen from among all the shapes in the document or all the shapes in the selection — to construct a shape range. For example, you could construct a  **ShapeRange** collection that contains the first three shapes in a document, all the selected shapes in a document, or all the freeform shapes in a document.


 **Note**  Most operations that you can do with a  **[Shape](http://msdn.microsoft.com/library/666cb7f0-62a8-f419-9838-007ef29506ee%28Office.15%29.aspx)** object, you can also do with a **ShapeRange** object that contains only one shape. Some operations, when performed on a **ShapeRange** object that contains more than one shape, will cause an error. This section describes how to:


- Return a set of shapes.
    
- Return a  **ShapeRange** object within a selection or range.
    
- Align, distribute, and group shapes in a  **ShapeRange** object.
    

## Example

Use  **Shapes.Range** (index), where index is the index number of the shape or an array that contains index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes in a publication. You can use Visual Basic's **Array** function to construct an array of index numbers. The following example sets the fill pattern for shapes one through three on the active publication.


```
Sub ChangeFillPattern() 
    ActiveDocument.Pages(1).Shapes.Range(Array(1, 2, 3)) _ 
        .Fill.PresetGradient Style:=msoGradientDiagonalDown, _ 
        Variant:=1, PresetGradientType:=msoGradientHorizon 
End Sub
```

Although you can use the  **[Range](http://msdn.microsoft.com/library/f9ef5314-21f1-378f-1552-fcd4e46f841d%28Office.15%29.aspx)** method to return any number of shapes, it is simpler to use the **[Item](http://msdn.microsoft.com/library/f316bbac-b0be-0281-585b-c32dcb709b66%28Office.15%29.aspx)** method if you want to return only a single member of the collection. For example, **Shapes** (1) is simpler than **Shapes.Range** (1).

Use  **Selection.ShapeRange** (index), where index is the index number of the shape, to return a **Shape** object that represents a shape within a selection. The following example selects the first two shapes on the first page of the active publication and then sets the fill for the first shape in the selection.




```
Sub ChangeFillForShapeRange() 
    ActiveDocument.Pages(1).Shapes.Range(Array(1, 2)).Select 
    Selection.ShapeRange(1).Fill.ForeColor.RGB = RGB(255, 0, 0) 
End Sub
```

This example selects all the shapes on the first page of the active publication, then adds and formats text in the second shape in the range.




```
Sub SelectShapesOnPageOne() 
    ActiveDocument.Pages(1).Shapes.Range.Select 
    With Selection.ShapeRange(2).TextFrame.TextRange 
        .Text = "Shape Number 2" 
        .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
        .Font.Size = 25 
    End With 
End Sub
```

Use the  **[Align](http://msdn.microsoft.com/library/ef522d47-3fc7-cfca-5b9a-44ff020f8b31%28Office.15%29.aspx)**, **[Distribute](http://msdn.microsoft.com/library/a145fb46-d7b6-bc3c-b7fd-cdb892fda179%28Office.15%29.aspx)**, or **[ZOrder](http://msdn.microsoft.com/library/2043f78c-ab83-e719-c3b5-5d75edcf1593%28Office.15%29.aspx)** method to position a set of shapes relative to each other or relative to the document. This example specifies a shape range and left-aligns and vertically distributes the shapes on the page.




```
Sub AlignDistibuteShapes() 
    Dim rngShapes As ShapeRange 
    Set rngShapes = ActiveDocument.Pages(1).Shapes.Range 
 
    With rngShapes 
        .Align AlignCmd:=msoAlignLefts, RelativeTo:=msoFalse 
        .Distribute DistributeCmd:=msoDistributeVertically, RelativeTo:=msoTrue 
    End With 
End Sub
```

Use the  **[Group](http://msdn.microsoft.com/library/ca3e011f-72ea-904e-da3f-cac7fe24341d%28Office.15%29.aspx)**, **[Regroup](http://msdn.microsoft.com/library/29342a78-9425-2356-963c-36a62a7f3091%28Office.15%29.aspx)**, or **[Ungroup](http://msdn.microsoft.com/library/253a366c-7317-14e7-2668-191eccec6cb8%28Office.15%29.aspx)** method to create and work with a single shape formed from a shape range. The **[GroupItems](http://msdn.microsoft.com/library/d37c75cd-a651-51d1-42c7-59879ccbbf1d%28Office.15%29.aspx)** property for a **Shape** object returns the **[GroupShapes](http://msdn.microsoft.com/library/dd723f99-25a9-81cc-1d16-eb7dcd651c5e%28Office.15%29.aspx)** object, which represents all the shapes that were grouped to form one shape. This example specifies a shape range and left-aligns and vertically distributes the shapes on the page.




```
Sub GroupShapes() 
    Dim rngShapes As ShapeRange 
    Set rngShapes = ActiveDocument.Pages(1).Shapes.Range 
    rngShapes.Group 
 
    rngShapes(1).Fill.OneColorGradient _ 
        Style:=msoGradientFromCenter, _ 
        Variant:=2, Degree:=1 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddToCatalogMergeArea](http://msdn.microsoft.com/library/6cb770c6-fe6e-ffe8-cd51-855d97b17aed%28Office.15%29.aspx)|
|[Align](http://msdn.microsoft.com/library/ef522d47-3fc7-cfca-5b9a-44ff020f8b31%28Office.15%29.aspx)|
|[Apply](http://msdn.microsoft.com/library/3531d0aa-479e-2d50-5e1e-a35f7c1e7ba6%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/11b9da00-85e4-fc7a-fa93-4a451b7bd15a%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/961d4646-8318-d2ff-ed98-649583d36115%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/fc9a7c2d-1bfc-d373-9d10-59df687b6fbf%28Office.15%29.aspx)|
|[Distribute](http://msdn.microsoft.com/library/a145fb46-d7b6-bc3c-b7fd-cdb892fda179%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/e940e551-4307-aa33-5713-80f77fade8af%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/fad24b08-9ada-0d6f-f526-ceec9ef996c1%28Office.15%29.aspx)|
|[GetHeight](http://msdn.microsoft.com/library/63501bf7-c24d-b58e-e4c5-c8a229f07c4e%28Office.15%29.aspx)|
|[GetLeft](http://msdn.microsoft.com/library/236717aa-368d-8403-5928-dc6c8e437c6f%28Office.15%29.aspx)|
|[GetTop](http://msdn.microsoft.com/library/bbee5dec-78fd-efd9-1368-2089a44d9bff%28Office.15%29.aspx)|
|[GetWidth](http://msdn.microsoft.com/library/a15d1b50-289a-8b02-e090-0f0a9637980a%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/ca3e011f-72ea-904e-da3f-cac7fe24341d%28Office.15%29.aspx)|
|[IncrementLeft](http://msdn.microsoft.com/library/1b760b5d-9879-5f64-c4c5-c9834a7928ff%28Office.15%29.aspx)|
|[IncrementRotation](http://msdn.microsoft.com/library/c58cdc12-948a-d6f8-2ddd-113008c7201b%28Office.15%29.aspx)|
|[IncrementTop](http://msdn.microsoft.com/library/8172406f-fac5-ad3d-49b8-cb4858d45c6d%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/f316bbac-b0be-0281-585b-c32dcb709b66%28Office.15%29.aspx)|
|[MoveIntoTextFlow](http://msdn.microsoft.com/library/bf76c82c-09de-5238-2c48-6addc5a4f000%28Office.15%29.aspx)|
|[MoveOutOfTextFlow](http://msdn.microsoft.com/library/36d6b22d-f041-6dd8-ce2c-9514ac6af5ae%28Office.15%29.aspx)|
|[PickUp](http://msdn.microsoft.com/library/ebd62b6e-807a-821c-d8ea-ed9be289c433%28Office.15%29.aspx)|
|[Regroup](http://msdn.microsoft.com/library/29342a78-9425-2356-963c-36a62a7f3091%28Office.15%29.aspx)|
|[RemoveFromCatalogMergeArea](http://msdn.microsoft.com/library/732cd277-9c2e-0a01-c2b5-8d016637884a%28Office.15%29.aspx)|
|[RerouteConnections](http://msdn.microsoft.com/library/ae2a64ea-1b7a-4ff6-304c-680dd96fd386%28Office.15%29.aspx)|
|[SaveAsBuildingBlock](http://msdn.microsoft.com/library/d68d5ccc-9f9f-4bc4-9748-37af9a6c3417%28Office.15%29.aspx)|
|[SaveAsPicture](http://msdn.microsoft.com/library/0be9b741-8f11-a386-313b-231a3269883a%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/54058fe5-d922-0ea9-08e8-99fff89bde55%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/8ff4eec9-9cf5-b6f0-062a-107aedbb8e38%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/3252ba74-d051-8c28-a9ed-c6f5ca711dec%28Office.15%29.aspx)|
|[SetShapesDefaultProperties](http://msdn.microsoft.com/library/1146cbf8-6d31-9fb8-c6a4-d54b68436cbd%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/253a366c-7317-14e7-2668-191eccec6cb8%28Office.15%29.aspx)|
|[ZOrder](http://msdn.microsoft.com/library/2043f78c-ab83-e719-c3b5-5d75edcf1593%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Adjustments](http://msdn.microsoft.com/library/819677e0-806d-a5ac-6fce-f7b0525e63ce%28Office.15%29.aspx)|
|[AlternativeText](http://msdn.microsoft.com/library/94cbb99b-3b35-76bb-e269-db8295b84f2f%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/6d594656-bd99-87ba-2244-fdba4ca471f4%28Office.15%29.aspx)|
|[AutoShapeType](http://msdn.microsoft.com/library/fa079239-07d8-0783-db34-77ee0f2d5391%28Office.15%29.aspx)|
|[BlackWhiteMode](http://msdn.microsoft.com/library/c85babbd-f05d-c3e1-3265-c08888eaf212%28Office.15%29.aspx)|
|[Callout](http://msdn.microsoft.com/library/25b9b444-6cbf-085a-df7f-8899e8e55057%28Office.15%29.aspx)|
|[ConnectionSiteCount](http://msdn.microsoft.com/library/f830739d-08be-562c-83fc-7f7a6f8e047c%28Office.15%29.aspx)|
|[Connector](http://msdn.microsoft.com/library/ce05006f-38b0-c04e-4a0f-dded72dfbc10%28Office.15%29.aspx)|
|[ConnectorFormat](http://msdn.microsoft.com/library/1a1516bd-ef27-0b37-09dd-45af8a531a76%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/5037bfe9-b430-4205-c514-b2f4313b4c53%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/cdff2b6f-52f5-3ab3-c57a-4647888cd96f%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/c9a479da-0b4e-9759-78ba-25006bd15ef9%28Office.15%29.aspx)|
|[GroupItems](http://msdn.microsoft.com/library/d37c75cd-a651-51d1-42c7-59879ccbbf1d%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/71ce4980-f5b5-c94c-c29d-32b97cf771fd%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/8a3b4f3b-3282-686b-f4fe-abf2d7677b3e%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/de6a638d-c197-a35b-130e-a9507d1b918e%28Office.15%29.aspx)|
|[HorizontalFlip](http://msdn.microsoft.com/library/c0dd2f4a-0baf-3720-113a-b929193f2b1d%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/34ec968c-af66-7629-066f-80c8e1b40e84%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/d7ad646b-be40-2ac4-9d3e-faa37f8bf456%28Office.15%29.aspx)|
|[InlineAlignment](http://msdn.microsoft.com/library/fed6d488-1483-2b59-b7be-1c4298f016a0%28Office.15%29.aspx)|
|[InlineTextRange](http://msdn.microsoft.com/library/5d7f3dfa-3e23-85c6-50cf-a6f960ccabfc%28Office.15%29.aspx)|
|[IsInline](http://msdn.microsoft.com/library/32e038cc-5837-93b4-de54-9bcd0549f1d4%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/2f071ef6-f6b3-2444-ea31-ea7abc9ef1ea%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/e9a6e8a0-f57a-63af-3040-5c43f8aba423%28Office.15%29.aspx)|
|[LinkFormat](http://msdn.microsoft.com/library/1f0add8d-7baa-65f0-e82b-a047a7bc0507%28Office.15%29.aspx)|
|[LockAspectRatio](http://msdn.microsoft.com/library/8ed4f41f-3395-dd59-29d4-f66afd19ac51%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/517eca4b-fa8c-0f6a-2829-75704bb4c899%28Office.15%29.aspx)|
|[Nodes](http://msdn.microsoft.com/library/513be66c-558c-f5f3-ed89-0ef4bc5a0101%28Office.15%29.aspx)|
|[OLEFormat](http://msdn.microsoft.com/library/237b51e8-dced-3e21-d257-410121107a63%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/3dd8c1bf-e204-422a-2719-12ace0550702%28Office.15%29.aspx)|
|[PictureFormat](http://msdn.microsoft.com/library/3d693c6b-b76b-0fe1-e7df-63fb08782f6f%28Office.15%29.aspx)|
|[Reflection](http://msdn.microsoft.com/library/db9ef973-39b9-7fe3-8b21-3ed1b74bb690%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/0239aaae-18c7-56ef-f2b1-82f82660370a%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/d6ee257c-9a26-abfc-9e8e-ef89bf627690%28Office.15%29.aspx)|
|[SoftEdge](http://msdn.microsoft.com/library/fd8006a9-91f8-6aeb-fa20-d5847122d14f%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/68221d37-505a-4701-8c9d-b8e695c8eb8f%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/792e3505-2c40-26e7-53c6-d50d84df22bb%28Office.15%29.aspx)|
|[TextEffect](http://msdn.microsoft.com/library/7bc822f2-4754-685d-fdd3-7479b5a3ac52%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/2dbb7fb4-3ae4-d4c1-8b7e-3e087e32a96f%28Office.15%29.aspx)|
|[TextWrap](http://msdn.microsoft.com/library/40fbc7aa-0a1b-7835-76bf-1815d7ccffc4%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/e5905f9d-dd84-b97e-ac5d-630f6c1208d7%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/9c4e6a86-2992-c0c8-6438-965e5c650dcf%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/31b56495-f3bb-73f4-52ef-eba4e43ea569%28Office.15%29.aspx)|
|[VerticalFlip](http://msdn.microsoft.com/library/cc3ab3ec-71f6-49fc-0141-505054d6abbb%28Office.15%29.aspx)|
|[Vertices](http://msdn.microsoft.com/library/0beb2323-8db6-c8c2-2f34-4c1ffde7fddc%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/8d512930-a908-f0e5-cd0d-dc8554acf0d5%28Office.15%29.aspx)|
|[Wizard](http://msdn.microsoft.com/library/43e24fbc-2dad-5fa6-9db8-a52ce86daab3%28Office.15%29.aspx)|
|[WizardTag](http://msdn.microsoft.com/library/49bdeff9-fec4-2b40-1650-cd78c9bce0d4%28Office.15%29.aspx)|
|[WizardTagInstance](http://msdn.microsoft.com/library/07d1c4c8-8efb-b029-2dba-37fef435cc8b%28Office.15%29.aspx)|
|[ZOrderPosition](http://msdn.microsoft.com/library/fc35f4dd-ef31-12e0-82a6-be2d0f765527%28Office.15%29.aspx)|

