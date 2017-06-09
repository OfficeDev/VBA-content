---
title: Shapes Object (Publisher)
keywords: vbapb10.chm2228223
f1_keywords:
- vbapb10.chm2228223
ms.prod: publisher
api_name:
- Publisher.Shapes
ms.assetid: 52e069a6-d54b-a11a-1cba-96174329cb02
ms.date: 06/08/2017
---


# Shapes Object (Publisher)

A collection of  **[Shape](http://msdn.microsoft.com/library/666cb7f0-62a8-f419-9838-007ef29506ee%28Office.15%29.aspx)** objects that represent all the shapes on a page of a publication. Each **Shape** object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


 **Note**  If you want to work with a subset of the shapes on a document — for example, to do something to only the AutoShapes on the document or to only the selected shapes — you must construct a  **[ShapeRange](shaperange-object-publisher.md)** collection that contains the shapes with which you want to work.


## Example

Use the  **[Shapes](http://msdn.microsoft.com/library/4e48d4cf-d7b6-9099-ddee-46a79e7eb7bf%28Office.15%29.aspx)** property to return the **Shapes** collection. The following example selects all the shapes on the first page of the active publication.


```
Sub SelectAllShapes() 
    ActiveDocument.Pages(1).Shapes.SelectAll 
End Sub
```


 **Note**  If you want to do something (like delete or set a property) to all the shapes in a publication at the same time, use the  **[Range](http://msdn.microsoft.com/library/f9ef5314-21f1-378f-1552-fcd4e46f841d%28Office.15%29.aspx)** method to create a **ShapeRange** object that contains all the shapes in the **Shapes** collection, and then apply the appropriate property or method to the **ShapeRange** object.



Use one of the following methods of the  **Shapes** collection: **[AddCallout](http://msdn.microsoft.com/library/bbf5f913-fcf0-b700-0c7e-9f0bdc7c6aea%28Office.15%29.aspx)**, **[AddConnector](http://msdn.microsoft.com/library/fd1ef969-7960-2555-e355-9804c86f6c01%28Office.15%29.aspx)**, **[AddCurve](http://msdn.microsoft.com/library/888a35cb-190d-4058-e0d7-a848d77ba920%28Office.15%29.aspx)**, **[AddLabel](http://msdn.microsoft.com/library/5a803aa2-d37f-6da1-7d8b-58ee2dcd8146%28Office.15%29.aspx)**, **[AddLine](http://msdn.microsoft.com/library/43df8878-5640-875f-06e0-37e1feb47b78%28Office.15%29.aspx)**, **[AddOLEObject](http://msdn.microsoft.com/library/c454f9cb-2005-5e55-80a7-6dfbe9c109e5%28Office.15%29.aspx)**, **[AddPolyline](http://msdn.microsoft.com/library/d49fb2bc-4df5-fff8-c741-2c0d35413fc5%28Office.15%29.aspx)**, **[AddShape](http://msdn.microsoft.com/library/500d8cb3-f066-fdb6-09ac-b03c7822e8bd%28Office.15%29.aspx)**, **[AddTextbox](http://msdn.microsoft.com/library/38494902-61d5-2017-819e-248b2b7bc0d1%28Office.15%29.aspx)**, or **[AddTextEffect](http://msdn.microsoft.com/library/21af82f1-d507-3c16-72df-bde1b5e00717%28Office.15%29.aspx)** to add a shape to a publication and return a **Shape** object that represents the newly created shape. The following example adds a new shape to the active publication.




```
Sub AddNewShape() 
    ActiveDocument.Pages(1).Shapes.AddShape Type:=msoShapeFoldedCorner, _ 
        Left:=50, Top:=50, Width:=100, Height:=200 
End Sub
```

Use  **Shapes** (index), where index is the index number, to return a single **Shape** object. The following example horizontally flips shape one on the first page of the active publication.




```
Sub FlipShape() 
    ActiveDocument.Pages(1).Shapes(1).Flip FlipCmd:=msoFlipHorizontal 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddBuildingBlock](http://msdn.microsoft.com/library/d875e97e-3519-4a88-916d-ec1a32654581%28Office.15%29.aspx)|
|[AddCallout](http://msdn.microsoft.com/library/bbf5f913-fcf0-b700-0c7e-9f0bdc7c6aea%28Office.15%29.aspx)|
|[AddCatalogMergeArea](http://msdn.microsoft.com/library/4af86b99-5a3a-b9f3-d269-16d635d35c83%28Office.15%29.aspx)|
|[AddCatalogMergeFieldToCanvas](http://msdn.microsoft.com/library/30cd45d0-97f0-ab01-31c2-8d819b435b1b%28Office.15%29.aspx)|
|[AddConnector](http://msdn.microsoft.com/library/fd1ef969-7960-2555-e355-9804c86f6c01%28Office.15%29.aspx)|
|[AddCurve](http://msdn.microsoft.com/library/888a35cb-190d-4058-e0d7-a848d77ba920%28Office.15%29.aspx)|
|[AddEmptyPictureFrame](http://msdn.microsoft.com/library/e473dea8-6d94-e9e4-ddb6-27c1fc8930e8%28Office.15%29.aspx)|
|[AddGroupWizard](http://msdn.microsoft.com/library/5a84f055-7f30-0757-f507-40ee34b214f4%28Office.15%29.aspx)|
|[AddLabel](http://msdn.microsoft.com/library/5a803aa2-d37f-6da1-7d8b-58ee2dcd8146%28Office.15%29.aspx)|
|[AddLine](http://msdn.microsoft.com/library/43df8878-5640-875f-06e0-37e1feb47b78%28Office.15%29.aspx)|
|[AddOLEObject](http://msdn.microsoft.com/library/c454f9cb-2005-5e55-80a7-6dfbe9c109e5%28Office.15%29.aspx)|
|[AddPicture](http://msdn.microsoft.com/library/a5305bd0-295f-46f6-7823-46dab750243b%28Office.15%29.aspx)|
|[AddPolyline](http://msdn.microsoft.com/library/d49fb2bc-4df5-fff8-c741-2c0d35413fc5%28Office.15%29.aspx)|
|[AddShape](http://msdn.microsoft.com/library/500d8cb3-f066-fdb6-09ac-b03c7822e8bd%28Office.15%29.aspx)|
|[AddTable](http://msdn.microsoft.com/library/1aa00f40-de41-12ed-8d4f-5e9c91cbf5af%28Office.15%29.aspx)|
|[AddTextbox](http://msdn.microsoft.com/library/38494902-61d5-2017-819e-248b2b7bc0d1%28Office.15%29.aspx)|
|[AddTextEffect](http://msdn.microsoft.com/library/21af82f1-d507-3c16-72df-bde1b5e00717%28Office.15%29.aspx)|
|[AddWebControl](http://msdn.microsoft.com/library/94b54939-9627-6b38-4375-f1c87fc8c4f7%28Office.15%29.aspx)|
|[AddWebNavigationBar](http://msdn.microsoft.com/library/26e9622c-ea28-b28b-9904-b3a3ccc9341b%28Office.15%29.aspx)|
|[AddWordArt](http://msdn.microsoft.com/library/8ff83baa-5d88-5f80-3a69-5f712ba5e583%28Office.15%29.aspx)|
|[BuildFreeform](http://msdn.microsoft.com/library/ea24a9a2-e72c-beb3-b17d-161ea41fff1d%28Office.15%29.aspx)|
|[FindShapeByWizardTag](http://msdn.microsoft.com/library/f1018f3a-4f8f-2686-ac58-6eee8827c743%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/174bbabb-e19f-4638-6dd8-780a8617fd70%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/435dd253-ae35-1dcf-ae5a-d7dfd40abf33%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/f9ef5314-21f1-378f-1552-fcd4e46f841d%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/67b88529-814d-c029-1bde-e5dade87636a%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/433bc241-b009-9d40-0630-5e81fbfc4064%28Office.15%29.aspx)|
|[CanvasArrangementType](http://msdn.microsoft.com/library/d86ee471-0c23-e6fc-d38c-b65e8c14d4c4%28Office.15%29.aspx)|
|[CanvasesCount](http://msdn.microsoft.com/library/d6755303-b05e-705f-bf15-cc6ec413c273%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/43052c93-461c-ca6a-3c8c-7142bd6d9ea1%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/dc05ea19-3c35-43ad-3ac8-f6402fce2011%28Office.15%29.aspx)|

