---
title: Shapes Object (PowerPoint)
keywords: vbapp10.chm543000
f1_keywords:
- vbapp10.chm543000
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes
ms.assetid: eb208855-254e-1a0f-884b-4a5edcfd584d
ms.date: 06/08/2017
---


# Shapes Object (PowerPoint)

A collection of all the  **[Shape](http://msdn.microsoft.com/library/1da93849-99e0-827e-ced3-c6cf7f8569f3%28Office.15%29.aspx)** objects on the specified slide.


## Remarks

Each  **Shape** object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


 **Note**  If you want to work with a subset of the shapes on a document — for example, to do something to only the AutoShapes on the document or to only the selected shapes — you must construct a  **[ShapeRange](shaperange-object-powerpoint.md)** collection that contains the shapes you want to work with. For an overview of how to work either with a single shape or with more than one shape at a time, see[How to: Work with Shapes (Drawing Objects)](http://msdn.microsoft.com/library/3ffaaaea-6406-262b-2bc7-788699175266%28Office.15%29.aspx).


## Example

Use the  **Shapes** property to return the **Shapes** collection. The following example selects all the shapes in the active presentation.


```
ActivePresentation.Slides(1).Shapes.SelectAll
```


 **Note**  If you want to do something (like delete or set a property) to all the shapes on a document at the same time, use the [Range](http://msdn.microsoft.com/library/5ee926d9-5b30-a26b-7365-f4709a1a7bdb%28Office.15%29.aspx)method with no argument to create a  **ShapeRange** object that contains all the shapes in the **Shapes** collection, and then apply the appropriate property or method to the **ShapeRange** object.

Use the [AddCallout](http://msdn.microsoft.com/library/e4b468d7-793a-09ae-fcfc-6a73db93c90e%28Office.15%29.aspx), [AddComment](http://msdn.microsoft.com/library/11347ca1-cef3-0923-2544-cb80e7fc5768%28Office.15%29.aspx), [AddConnector](http://msdn.microsoft.com/library/407eee86-11c1-7bee-ed25-aba71a930a1c%28Office.15%29.aspx), [AddCurve](http://msdn.microsoft.com/library/47f90182-a71b-a028-c43f-a85d59d2a56b%28Office.15%29.aspx), [AddLabel](http://msdn.microsoft.com/library/b744daf1-5b99-9649-8b97-d3f2193373c1%28Office.15%29.aspx), [AddLine](http://msdn.microsoft.com/library/9dbe640b-5ba4-a620-d3c6-4a2d0cc2bc27%28Office.15%29.aspx), [AddMediaObject](http://msdn.microsoft.com/library/7e2ab704-7fd4-86d7-3f61-8d69c13b5685%28Office.15%29.aspx), [AddOLEObject](http://msdn.microsoft.com/library/88a5aa63-0531-b9d8-43d2-5a995b91b189%28Office.15%29.aspx), [AddPicture](http://msdn.microsoft.com/library/af432432-b09b-3ca6-d392-132bd78251c7%28Office.15%29.aspx), [AddPlaceholder](http://msdn.microsoft.com/library/10927d59-1810-2f91-eb52-c42113570ccc%28Office.15%29.aspx), [AddPolyline](http://msdn.microsoft.com/library/e42c4f7a-de68-88bf-d250-28e642b56232%28Office.15%29.aspx), [AddShape](http://msdn.microsoft.com/library/2bc6cce5-3461-61ff-083d-bd36ee71cb59%28Office.15%29.aspx), [AddTable](http://msdn.microsoft.com/library/77ce193e-10f7-25f4-a6f8-99d7d2b781ad%28Office.15%29.aspx), [AddTextbox](http://msdn.microsoft.com/library/0c7c6093-48f6-e1f1-1837-e69d6ef13e57%28Office.15%29.aspx), [AddTextEffect](http://msdn.microsoft.com/library/4428ac57-c704-475a-1640-78a556e9ac3d%28Office.15%29.aspx), or [AddTitle](http://msdn.microsoft.com/library/1fe13529-526a-1b29-7589-c155f9e46379%28Office.15%29.aspx)method to create a new shape and add it to the  **Shapes** collection. Use the[BuildFreeform](http://msdn.microsoft.com/library/330ea348-9f8c-c418-d67f-e4fd6c105c59%28Office.15%29.aspx)method in conjunction with the [ConvertToShape](http://msdn.microsoft.com/library/bc3d209e-6735-3011-9334-46049d269355%28Office.15%29.aspx)method to create a new freeform and add it to the collection. The following example adds a rectangle to the active presentation.




```
ActivePresentation.Slides(1).Shapes.AddShape Type:=msoShapeRectangle, _

    Left:=50, Top:=50, Width:=100, Height:=200
```

Use  **Shapes** (index), where index is the shape's name or index number, to return a single **Shape** object. The following example sets the fill to a preset shade for shape one in the active presentation.




```
ActivePresentation.Slides(1).Shapes(1).Fill _

    .PresetGradient Style:=msoGradientHorizontal, Variant:=1, _

    PresetGradientType:=msoGradientBrass
```

Use  **Shapes.Range** (index), where index is the shape's name or index number or an array of shape names or index numbers, to return a **[ShapeRange](shaperange-object-powerpoint.md)** collection that represents a subset of the **Shapes** collection. The following example sets the fill pattern for shapes one and three in the active presentation.




```
ActivePresentation.Slides(1).Shapes.Range(Array(1, 3)).Fill _

    .Patterned Pattern:=msoPatternHorizontalBrick
```

Use  **Shapes.Placeholders** (index), where index is the placeholder number, to return a **Shape** object that represents a placeholder. If the specified slide has a title, use **Shapes.Placeholders(1)** or **Shapes.Title** to return the title placeholder. The following example adds a slide to the active presentation and then adds text to both the title and the subtitle (the subtitle is the second placeholder on a slide with this layout).




```
With ActivePresentation.Slides.Add(Index:=1, Layout:=ppLayoutTitle).Shapes

    .Title.TextFrame.TextRange = "This is the title text"

    .Placeholders(2).TextFrame.TextRange = "This is subtitle text"

End With
```


## Methods



|**Name**|
|:-----|
|[AddCallout](http://msdn.microsoft.com/library/e4b468d7-793a-09ae-fcfc-6a73db93c90e%28Office.15%29.aspx)|
|[AddChart2](http://msdn.microsoft.com/library/07f225bc-1c0d-cca5-b6a3-9de0a018eb4c%28Office.15%29.aspx)|
|[AddConnector](http://msdn.microsoft.com/library/407eee86-11c1-7bee-ed25-aba71a930a1c%28Office.15%29.aspx)|
|[AddCurve](http://msdn.microsoft.com/library/47f90182-a71b-a028-c43f-a85d59d2a56b%28Office.15%29.aspx)|
|[AddInkShapeFromXML](http://msdn.microsoft.com/library/88a395ac-b11e-d42e-f4b4-b41bf1d1347e%28Office.15%29.aspx)|
|[AddLabel](http://msdn.microsoft.com/library/b744daf1-5b99-9649-8b97-d3f2193373c1%28Office.15%29.aspx)|
|[AddLine](http://msdn.microsoft.com/library/9dbe640b-5ba4-a620-d3c6-4a2d0cc2bc27%28Office.15%29.aspx)|
|[AddMediaObject2](http://msdn.microsoft.com/library/157499e5-1b90-d85f-b1d8-85a115fc907e%28Office.15%29.aspx)|
|[AddMediaObjectFromEmbedTag](http://msdn.microsoft.com/library/c463e7e2-8bac-8762-fec8-e1e84847907b%28Office.15%29.aspx)|
|[AddOLEObject](http://msdn.microsoft.com/library/88a5aa63-0531-b9d8-43d2-5a995b91b189%28Office.15%29.aspx)|
|[AddPicture](http://msdn.microsoft.com/library/af432432-b09b-3ca6-d392-132bd78251c7%28Office.15%29.aspx)|
|[AddPicture2](http://msdn.microsoft.com/library/2956fa14-40bb-458a-aef1-caceab15e067%28Office.15%29.aspx)|
|[AddPlaceholder](http://msdn.microsoft.com/library/10927d59-1810-2f91-eb52-c42113570ccc%28Office.15%29.aspx)|
|[AddPolyline](http://msdn.microsoft.com/library/e42c4f7a-de68-88bf-d250-28e642b56232%28Office.15%29.aspx)|
|[AddShape](http://msdn.microsoft.com/library/2bc6cce5-3461-61ff-083d-bd36ee71cb59%28Office.15%29.aspx)|
|[AddSmartArt](http://msdn.microsoft.com/library/5bd66a76-a31c-3633-7aae-f24e0a92021c%28Office.15%29.aspx)|
|[AddTable](http://msdn.microsoft.com/library/77ce193e-10f7-25f4-a6f8-99d7d2b781ad%28Office.15%29.aspx)|
|[AddTextbox](http://msdn.microsoft.com/library/0c7c6093-48f6-e1f1-1837-e69d6ef13e57%28Office.15%29.aspx)|
|[AddTextEffect](http://msdn.microsoft.com/library/4428ac57-c704-475a-1640-78a556e9ac3d%28Office.15%29.aspx)|
|[AddTitle](http://msdn.microsoft.com/library/1fe13529-526a-1b29-7589-c155f9e46379%28Office.15%29.aspx)|
|[BuildFreeform](http://msdn.microsoft.com/library/330ea348-9f8c-c418-d67f-e4fd6c105c59%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/f6c5eac1-3b65-3023-3b7a-557c7bfb0f02%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/8aa534f8-bd59-3945-cc1f-45ffc3883bf7%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/6a1e5b6d-da09-fae8-7165-0c9bf71d525c%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/5ee926d9-5b30-a26b-7365-f4709a1a7bdb%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/9d3f5b93-2a8b-5b9a-d725-729baa190a38%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/23c2ea6f-ed51-4a1a-0e00-94f891242c0a%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/bc313541-1e87-cc85-e489-80d53f18abe5%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/495a5a34-efdb-784e-8748-7bc6005e7ffd%28Office.15%29.aspx)|
|[HasTitle](http://msdn.microsoft.com/library/0754bda8-7e19-6dd1-55a3-2b19541480b9%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/b6d9ba88-0073-3482-b7fb-5f9d36f79b48%28Office.15%29.aspx)|
|[Placeholders](http://msdn.microsoft.com/library/2926d893-056a-0805-85ba-681e64bf81ed%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/61e5f162-d9dd-f8d3-6c15-d5a40c00c10f%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
