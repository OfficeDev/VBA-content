---
title: Shape Object (Publisher)
keywords: vbapb10.chm2293759
f1_keywords:
- vbapb10.chm2293759
ms.prod: publisher
api_name:
- Publisher.Shape
ms.assetid: 666cb7f0-62a8-f419-9838-007ef29506ee
ms.date: 06/08/2017
---


# Shape Object (Publisher)

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, ActiveX control, or picture. The  **Shape** object is a member of the **[Shapes](http://msdn.microsoft.com/library/52e069a6-d54b-a11a-1cba-96174329cb02%28Office.15%29.aspx)** collection, which includes all the shapes on a page or in a selection.


 **Note**  There are three objects that represent shapes: the  **Shapes** collection, which represents all the shapes on a document; the **[ShapeRange](http://msdn.microsoft.com/library/c85967c9-af43-747d-7e0b-64ddc22c84be%28Office.15%29.aspx)** collection, which represents a specified subset of the shapes on a document (for example, a **ShapeRange** object could represent shapes one and four on the document, or it could represent all the selected shapes on the document); the **Shape** object, which represents a single shape on a document. If you want to work with several shape at the same time or with shapes within the selection, use a **ShapeRange** collection. This section describes how to:


- Return an existing shape on a document.
    
- Return a shape or shapes within a selection.
    
- Return a newly created shape.
    
- Work with a group of shapes.
    
- Format a shape.
    
- Use other important shape properties.
    

## Example

Use  **[Shapes](http://msdn.microsoft.com/library/52e069a6-d54b-a11a-1cba-96174329cb02%28Office.15%29.aspx)** (index), where index is the name or the index number, to return a single **Shape** object. The following example horizontally flips shape one on the active document.


```
Sub FlipShape() 
    ActiveDocument.Pages(1).Shapes(1).Flip FlipCmd:=msoFlipHorizontal 
End Sub
```

The following example horizontally flips the shape named "Rectangle 1" on the active document.




```
Sub FlipShapeByName() 
    ActiveDocument.Pages(1).Shapes("Rectangle 1") _ 
        .Flip FlipCmd:=msoFlipHorizontal 
End Sub
```

Each shape is assigned a default name when it is created. For example, if you add three different shapes to a document, they might be named "Rectangle 2," "TextBox 3," and "Oval 4." To give a shape a more meaningful name, set the  **Name** property of the shape.

Use  **Selection.ShapeRange** (index), where index is the name or the index number, to return a **Shape** object that represents a shape within a selection. The following example sets the fill for the first shape in the selection, assuming that the selection contains at least one shape.




```
Sub FillSelectedShape() 
    Selection.ShapeRange(1).Fill.ForeColor.RGB = RGB(255, 0, 0) 
End Sub
```

The following example sets the fill for all the shapes in the selection, assuming that the selection contains at least one shape.




```
Sub FillAllSelectedShapes() 
    Dim shpShape As Shape 
    For Each
```




```
shpShape In Selection.ShapeRange 
       
```




```
shpShape.Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
    Next shpShape 
End Sub
```

To add a  **Shape** object to the collection of shapes for the specified document and return a **Shape** object that represents the newly created shape, use one of the following methods of the **Shapes** collection: **[AddCallout](http://msdn.microsoft.com/library/bbf5f913-fcf0-b700-0c7e-9f0bdc7c6aea%28Office.15%29.aspx)**, **[AddConnector](http://msdn.microsoft.com/library/fd1ef969-7960-2555-e355-9804c86f6c01%28Office.15%29.aspx)**, **[AddCurve](http://msdn.microsoft.com/library/888a35cb-190d-4058-e0d7-a848d77ba920%28Office.15%29.aspx)**, **[AddLabel](http://msdn.microsoft.com/library/5a803aa2-d37f-6da1-7d8b-58ee2dcd8146%28Office.15%29.aspx)**, **[AddLine](http://msdn.microsoft.com/library/43df8878-5640-875f-06e0-37e1feb47b78%28Office.15%29.aspx)**, **[AddOLEObject](http://msdn.microsoft.com/library/c454f9cb-2005-5e55-80a7-6dfbe9c109e5%28Office.15%29.aspx)**, **[AddPolyline](http://msdn.microsoft.com/library/d49fb2bc-4df5-fff8-c741-2c0d35413fc5%28Office.15%29.aspx)**, **[AddShape](http://msdn.microsoft.com/library/500d8cb3-f066-fdb6-09ac-b03c7822e8bd%28Office.15%29.aspx)**, **[AddTextBox](http://msdn.microsoft.com/library/38494902-61d5-2017-819e-248b2b7bc0d1%28Office.15%29.aspx)** or **[AddTextEffect](http://msdn.microsoft.com/library/21af82f1-d507-3c16-72df-bde1b5e00717%28Office.15%29.aspx)**. The following example adds a rectangle to the active document.




```
Sub AddNewShape() 
    ActiveDocument.Pages(1).Shapes.AddShape Type:=msoShapeRectangle, _ 
        Left:=400, Top:=72, Width:=100, Height:=200 
End Sub
```

Use  **[GroupItems](http://msdn.microsoft.com/library/9194f43b-bd8a-76a9-aa8c-17544d052d47%28Office.15%29.aspx)** (index), where index is the shape name or the index number within the group, to return a **Shape** object that represents a single shape in a grouped shape. Use the **[Group](http://msdn.microsoft.com/library/ca3e011f-72ea-904e-da3f-cac7fe24341d%28Office.15%29.aspx)** or **[Regroup](http://msdn.microsoft.com/library/29342a78-9425-2356-963c-36a62a7f3091%28Office.15%29.aspx)** method to group a range of shapes and return a single **Shape** object that represents the newly formed group. After a group has been formed, you can work with the group the same way you work with any other shape. This example adds three shapes to the active publication, groups the shapes, and sets the fill color for each of the shapes in the group




```
Sub WorkWithGroupShapes() 
 
    With ActiveDocument.Pages(1).Shapes 
        .AddShape Type:=msoShapeIsoscelesTriangle, Left:=100, _ 
            Top:=72, Width:=100, Height:=100 
        .AddShape Type:=msoShapeIsoscelesTriangle, Left:=250, _ 
            Top:=72, Width:=100, Height:=100 
        .AddShape Type:=msoShapeIsoscelesTriangle, Left:=400, _ 
            Top:=72, Width:=100, Height:=100 
        .SelectAll 
 
        With Selection.ShapeRange 
            .Group 
            .GroupItems(1).Fill.ForeColor _ 
                .RGB = RGB(Red:=255, Green:=0, Blue:=0) 
            .GroupItems(2).Fill.ForeColor _ 
                .RGB = RGB(Red:=0, Green:=255, Blue:=0) 
            .GroupItems(3).Fill.ForeColor _ 
                .RGB = RGB(Red:=0, Green:=0, Blue:=255) 
        End With 
    End With 
End Sub
```

Use the  **[Fill](http://msdn.microsoft.com/library/ff1b8d02-150e-e023-2f0a-b1608cc99644%28Office.15%29.aspx)** property to return the **[FillFormat](http://msdn.microsoft.com/library/0a5d4f7a-c42a-28ad-c86d-ac9828a3b874%28Office.15%29.aspx)** object, which contains all the properties and methods for formatting the fill of a closed shape. The **[Shadow](http://msdn.microsoft.com/library/cfb908ae-ef1d-9539-1f82-2693cbe38d97%28Office.15%29.aspx)** property returns the **[ShadowFormat](http://msdn.microsoft.com/library/b23ab92e-5e49-8d8d-69d5-93d391a9edb2%28Office.15%29.aspx)** object, which you use to format a shadow. Use the **[Line](http://msdn.microsoft.com/library/3d53f917-87ad-159d-65c3-e6fdfa72b15e%28Office.15%29.aspx)** property to return a **[LineFormat](http://msdn.microsoft.com/library/9c973f5a-b2d2-78b1-24c3-350f1ba4c2ab%28Office.15%29.aspx)** object, which contains properties and methods for formatting lines and arrows. The **[TextEffect](http://msdn.microsoft.com/library/187b55f8-9593-6a00-61e6-dbcf5c56b987%28Office.15%29.aspx)** property returns the **[TextEffectFormat](http://msdn.microsoft.com/library/672d0ef0-cbcd-05ef-9aa5-b986c7b045ac%28Office.15%29.aspx)** object, which you use to format WordArt. The **[Callout](http://msdn.microsoft.com/library/e0682bb4-1129-fa58-b28c-46d7ce2fad0c%28Office.15%29.aspx)** property returns the **[CalloutFormat](http://msdn.microsoft.com/library/1f54aba3-3872-e668-fe76-1966d1a62cca%28Office.15%29.aspx)** object, which you use to format line callouts. The **[TextWrap](http://msdn.microsoft.com/library/e641d9a5-5b63-06d0-a0c3-d3feb1910159%28Office.15%29.aspx)** property returns the **[WrapFormat](http://msdn.microsoft.com/library/b6f80d40-2043-6944-3ed8-f26635c7fa4d%28Office.15%29.aspx)** object, which you use to define how text wraps around shapes. The **[ThreeD](http://msdn.microsoft.com/library/e3430bb2-2f2a-14a6-8eb4-98a29a96ad1c%28Office.15%29.aspx)** property returns the **[ThreeDFormat](http://msdn.microsoft.com/library/11d57330-c99e-5aa9-d47c-2c5d2846ed4d%28Office.15%29.aspx)** object, which you use to create 3-D shapes. You can use the **[PickUp](http://msdn.microsoft.com/library/12b59235-db2d-b451-de8e-9e8df6bfeb1c%28Office.15%29.aspx)** and **[Apply](http://msdn.microsoft.com/library/711c72b6-3618-be0b-fb72-9f68fdbcc4a8%28Office.15%29.aspx)** methods to transfer formatting from one shape to another.



Use the  **[SetShapesDefaultProperties](http://msdn.microsoft.com/library/3f7d7143-3a08-6ff4-c28e-86598212a876%28Office.15%29.aspx)** method for a **Shape** object to set the formatting for the default shape for the document. New shapes inherit many of their attributes from the default shape.

Use the  **[Type](http://msdn.microsoft.com/library/bb712dd4-5d81-10e0-9b4c-4af6a09a3c71%28Office.15%29.aspx)** property to specify the type of shape: freeform, AutoShape, OLE object, callout, or linked picture, for instance. Use the **[AutoShapeType](http://msdn.microsoft.com/library/f469dc31-a620-5561-ce57-fbff8a5536c0%28Office.15%29.aspx)** property to specify the type of AutoShape: oval, rectangle, or balloon, for instance.



Use the  **[Width](http://msdn.microsoft.com/library/0b7c5b57-1968-dabb-1e19-9f1d450cea7f%28Office.15%29.aspx)** and **[Height](http://msdn.microsoft.com/library/2796ae7e-f4b9-4d79-ff98-d5807286b41e%28Office.15%29.aspx)** properties to specify the size of the shape.



Use  **[TextFrame](http://msdn.microsoft.com/library/fc654905-d56b-9a6c-28fa-4b54bf2a8686%28Office.15%29.aspx)** and **[TextRange](http://msdn.microsoft.com/library/31aa92d1-852f-3742-defa-94485411bcc3%28Office.15%29.aspx)** properties to return the **[TextFrame](http://msdn.microsoft.com/library/95e88f5a-b3dc-272e-7c1d-5282c97ae11e%28Office.15%29.aspx)** and **[TextRange](http://msdn.microsoft.com/library/566f240b-d2a6-8cb3-9eb7-68328d6c28bd%28Office.15%29.aspx)** objects, respectively, which contain all the properties and methods for inserting and formatting text within shapes and publications and linking the text frames together. The following example adds a text box to the first page of the active publication, then adds text to it and formats the text.




```
Sub CreateNewTextBox() 
    With ActiveDocument.Pages(1).Shapes.AddTextbox( _ 
        Orientation:=pbTextOrientationHorizontal, Left:=100, _ 
        Top:=100, Width:=200, Height:=100).TextFrame.TextRange 
        .Text = "This is a textbox." 
        With .Font 
            .Name = "Stencil" 
            .Bold = msoTrue 
            .Size = 30 
        End With 
    End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddToCatalogMergeArea](http://msdn.microsoft.com/library/4178d286-045f-a7b6-86b6-710bed10e824%28Office.15%29.aspx)|
|[Apply](http://msdn.microsoft.com/library/711c72b6-3618-be0b-fb72-9f68fdbcc4a8%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/cfec06d8-9f9b-4d88-eb28-e9e29fb1aed1%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/d800c1e5-7655-9071-a373-7772fa1ca15f%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/29dc0685-b354-427c-2b95-e02847dbb09e%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/9f35a496-5312-bff1-a31e-05baaaf69e92%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/6d0004a5-2d76-955a-64ff-140dfbc313f3%28Office.15%29.aspx)|
|[GetHeight](http://msdn.microsoft.com/library/e94eaede-f2b3-4f68-b3ec-915354a1b0b7%28Office.15%29.aspx)|
|[GetLeft](http://msdn.microsoft.com/library/e8f28ab3-f9da-eae7-2a21-b8b2505e9b44%28Office.15%29.aspx)|
|[GetTop](http://msdn.microsoft.com/library/65421a42-a16a-2c9d-c510-f1c6066ae0bb%28Office.15%29.aspx)|
|[GetWidth](http://msdn.microsoft.com/library/9df33329-c37b-82f5-93b4-fc4752ee907e%28Office.15%29.aspx)|
|[IncrementLeft](http://msdn.microsoft.com/library/447886ad-f515-9869-524a-a803ab025fa4%28Office.15%29.aspx)|
|[IncrementRotation](http://msdn.microsoft.com/library/3293c707-f3e8-1afb-cf9c-231ceae66ab6%28Office.15%29.aspx)|
|[IncrementTop](http://msdn.microsoft.com/library/c7a5bf47-7c5a-f6e8-b2b7-c95bea9dc081%28Office.15%29.aspx)|
|[MoveIntoTextFlow](http://msdn.microsoft.com/library/d8a2af57-f974-717e-0d97-c8a3aee16f01%28Office.15%29.aspx)|
|[MoveOutOfTextFlow](http://msdn.microsoft.com/library/44411d6b-a627-f0c1-0576-2918f586ff0b%28Office.15%29.aspx)|
|[MoveToPage](http://msdn.microsoft.com/library/1893035f-6739-7480-6ba0-2ca6a42355fa%28Office.15%29.aspx)|
|[PickUp](http://msdn.microsoft.com/library/12b59235-db2d-b451-de8e-9e8df6bfeb1c%28Office.15%29.aspx)|
|[RemoveCatalogMergeArea](http://msdn.microsoft.com/library/addff960-562e-b8e8-ec56-ddcf2b9ccaa7%28Office.15%29.aspx)|
|[RemoveFromCatalogMergeArea](http://msdn.microsoft.com/library/3b3630c3-6bf1-494b-151c-c930f32a2a77%28Office.15%29.aspx)|
|[RerouteConnections](http://msdn.microsoft.com/library/04afd4aa-dc84-d39c-e9fa-d06f8f4c0a02%28Office.15%29.aspx)|
|[SaveAsBuildingBlock](http://msdn.microsoft.com/library/5dd51d12-9bb2-4dd5-9b4c-20f755beef12%28Office.15%29.aspx)|
|[SaveAsPicture](http://msdn.microsoft.com/library/2cc18a83-b947-ca8c-eab4-71a03b79b82b%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/733afebc-0946-07eb-0550-547a4dc9f9da%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/07dcc04e-cb84-9c69-c589-87c0ff0bb147%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/d18914fd-7679-e922-090c-78affdb39d6a%28Office.15%29.aspx)|
|[SetCaption](http://msdn.microsoft.com/library/dd3ca08b-06c7-4a12-b51c-5d76ce1601b5%28Office.15%29.aspx)|
|[SetShapesDefaultProperties](http://msdn.microsoft.com/library/3f7d7143-3a08-6ff4-c28e-86598212a876%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/2edd16fc-d607-856f-0524-bdef1e58a9da%28Office.15%29.aspx)|
|[ZOrder](http://msdn.microsoft.com/library/05143a2b-924e-b5a3-390d-9493627bfa9f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Adjustments](http://msdn.microsoft.com/library/14794cba-c671-51e3-0aac-52e885a4ba7f%28Office.15%29.aspx)|
|[AlternativeText](http://msdn.microsoft.com/library/13bc57af-7067-d60c-5096-a68b1f821d58%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/d39e9ba7-9e08-a903-8c44-ede0174ad2f4%28Office.15%29.aspx)|
|[AutoShapeType](http://msdn.microsoft.com/library/f469dc31-a620-5561-ce57-fbff8a5536c0%28Office.15%29.aspx)|
|[BlackWhiteMode](http://msdn.microsoft.com/library/0a735488-956f-bd3c-ad74-1639780e4e24%28Office.15%29.aspx)|
|[BorderArt](http://msdn.microsoft.com/library/dcc0ceb4-ef69-ffd3-e510-13dcb8d06832%28Office.15%29.aspx)|
|[Callout](http://msdn.microsoft.com/library/e0682bb4-1129-fa58-b28c-46d7ce2fad0c%28Office.15%29.aspx)|
|[CatalogMergeItems](http://msdn.microsoft.com/library/1dcf4ae0-7a18-f1d5-2176-1912c63eefcc%28Office.15%29.aspx)|
|[ConnectionSiteCount](http://msdn.microsoft.com/library/00c32910-96b6-6981-8359-de4a71852934%28Office.15%29.aspx)|
|[Connector](http://msdn.microsoft.com/library/6cdff1e7-59b0-9905-96f8-99b79db1acd5%28Office.15%29.aspx)|
|[ConnectorFormat](http://msdn.microsoft.com/library/280c424c-530c-55ab-da4f-65b858ee3dd8%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/ff1b8d02-150e-e023-2f0a-b1608cc99644%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/f138e966-4b01-8cd2-36e7-d9d10b33062f%28Office.15%29.aspx)|
|[GroupItems](http://msdn.microsoft.com/library/9194f43b-bd8a-76a9-aa8c-17544d052d47%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/6f544d9c-00a4-3047-fbfb-6f1835bbe2c6%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/faf9514a-438b-ad12-a830-ed34cea8ba03%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/2796ae7e-f4b9-4d79-ff98-d5807286b41e%28Office.15%29.aspx)|
|[HorizontalFlip](http://msdn.microsoft.com/library/5a940631-c63a-efdf-6cfb-dc6b82594028%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/0990ab32-b4a3-6c89-cb9f-8f8c64ef804f%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/df4ccd93-e3fa-eeef-b5ea-e99aa0dde199%28Office.15%29.aspx)|
|[InlineAlignment](http://msdn.microsoft.com/library/daef2761-2a93-25da-9c12-1fed0fdd24ab%28Office.15%29.aspx)|
|[InlineTextRange](http://msdn.microsoft.com/library/40b0ea73-499d-a930-da09-2f20066b7129%28Office.15%29.aspx)|
|[IsExcess](http://msdn.microsoft.com/library/217689d6-7508-92ab-3828-e61fc70f0993%28Office.15%29.aspx)|
|[IsGroupMember](http://msdn.microsoft.com/library/bbd9b662-b47d-d5cf-6858-e208c44f88a0%28Office.15%29.aspx)|
|[IsInline](http://msdn.microsoft.com/library/5c5c6181-070f-2a66-8d70-2d6372cb365e%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/275f5af9-9812-2a6b-bba3-704d4a7f5601%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/3d53f917-87ad-159d-65c3-e6fdfa72b15e%28Office.15%29.aspx)|
|[LinkFormat](http://msdn.microsoft.com/library/801c3a87-7cc6-8c7b-094a-55e8d8d7a004%28Office.15%29.aspx)|
|[LockAspectRatio](http://msdn.microsoft.com/library/eeb87bb5-01d5-5d21-b268-045497ea3682%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/307c131b-f6ad-38e7-d214-420063d3e5ec%28Office.15%29.aspx)|
|[Nodes](http://msdn.microsoft.com/library/a1463ff3-5b75-e4b9-df12-985538713c7c%28Office.15%29.aspx)|
|[OLEFormat](http://msdn.microsoft.com/library/36bffb6b-4c7b-85f9-87b3-d7d7c1aed134%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/579a8ddf-5da6-905a-2784-f9083d4a1ad6%28Office.15%29.aspx)|
|[ParentGroupShape](http://msdn.microsoft.com/library/ced4c348-4ef5-c703-fdea-65c33d37b4c0%28Office.15%29.aspx)|
|[PictureFormat](http://msdn.microsoft.com/library/2a812ba3-18e4-fc42-6d07-535511a79650%28Office.15%29.aspx)|
|[Reflection](http://msdn.microsoft.com/library/a9a12d07-8edc-2f1b-9f7d-4aeae43b1335%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/3cb55e8c-83fa-2f20-caac-a1e897e9a369%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/cfb908ae-ef1d-9539-1f82-2693cbe38d97%28Office.15%29.aspx)|
|[SoftEdge](http://msdn.microsoft.com/library/1bbb441e-314d-30d6-bae7-f96f81224dd9%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/a9b29d1f-2459-556c-56f8-f8f809b879c9%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/282f77c8-f075-1eeb-65e8-f1126def32ff%28Office.15%29.aspx)|
|[TextEffect](http://msdn.microsoft.com/library/187b55f8-9593-6a00-61e6-dbcf5c56b987%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/fc654905-d56b-9a6c-28fa-4b54bf2a8686%28Office.15%29.aspx)|
|[TextWrap](http://msdn.microsoft.com/library/e641d9a5-5b63-06d0-a0c3-d3feb1910159%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/e3430bb2-2f2a-14a6-8eb4-98a29a96ad1c%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/76ab84a9-651c-ddc6-6f7f-f98e2b71074f%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/bb712dd4-5d81-10e0-9b4c-4af6a09a3c71%28Office.15%29.aspx)|
|[VerticalFlip](http://msdn.microsoft.com/library/b3c7492f-08ee-8fad-102a-8e2a2f69b969%28Office.15%29.aspx)|
|[Vertices](http://msdn.microsoft.com/library/40b4800f-b17c-eff4-cb87-1e2d44d53ee3%28Office.15%29.aspx)|
|[WebCheckBox](http://msdn.microsoft.com/library/13796525-584f-7109-5dea-1f2baf1efda7%28Office.15%29.aspx)|
|[WebCommandButton](http://msdn.microsoft.com/library/c20b937b-6f53-fdc1-830a-4044831c351a%28Office.15%29.aspx)|
|[WebListBox](http://msdn.microsoft.com/library/c100dfc7-6fbd-db48-4de9-4a9a49739a8f%28Office.15%29.aspx)|
|[WebNavigationBarSetName](http://msdn.microsoft.com/library/0d9abe17-6936-562b-9210-5f092d13f215%28Office.15%29.aspx)|
|[WebOptionButton](http://msdn.microsoft.com/library/0c43387c-0cb6-5d6f-68cb-d1883ce17243%28Office.15%29.aspx)|
|[WebTextBox](http://msdn.microsoft.com/library/8a3f8389-728f-b8ae-3c89-dc8d03a3818e%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/0b7c5b57-1968-dabb-1e19-9f1d450cea7f%28Office.15%29.aspx)|
|[Wizard](http://msdn.microsoft.com/library/89014daf-66dc-7913-0b0e-ac80f6e85791%28Office.15%29.aspx)|
|[WizardTag](http://msdn.microsoft.com/library/b93bbdf9-6ce7-3ba6-566a-b11f8044fbda%28Office.15%29.aspx)|
|[WizardTagInstance](http://msdn.microsoft.com/library/908d3f31-f277-7213-737e-9a946687bda7%28Office.15%29.aspx)|
|[ZOrderPosition](http://msdn.microsoft.com/library/46eb765b-578e-f6df-43b7-c14443cddbb2%28Office.15%29.aspx)|

