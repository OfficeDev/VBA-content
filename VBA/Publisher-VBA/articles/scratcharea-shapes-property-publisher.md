---
title: ScratchArea.Shapes Property (Publisher)
keywords: vbapb10.chm1179651
f1_keywords:
- vbapb10.chm1179651
ms.prod: publisher
api_name:
- Publisher.ScratchArea.Shapes
ms.assetid: 0d867fec-42f4-fd61-c6c3-745be955e5d2
ms.date: 06/08/2017
---


# ScratchArea.Shapes Property (Publisher)

Returns a  **[Shapes](shapes-object-publisher.md)** collection that represents all the  **Shape** objects in the specified publication. This collection can contain drawings, shapes, pictures, OLE objects, ActiveX controls, text objects, and callouts.


## Syntax

 _expression_. **Shapes**

 _expression_A variable that represents a  **ScratchArea** object.


## Remarks

For information about returning a single member of a collection, see  **Returning an Object from a Collection**.


## Example

This example adds a rectangle to the first page in the active publication.


```vb
Sub AddNewRectangle() 
 ActiveDocument.Pages(1).Shapes.AddShape Type:=msoShapeRectangle, _ 
 Left:=5, Top:=25, Width:=100, Height:=50 
End Sub
```

This example sets the fill texture for all the shapes in the active publication. This example assumes there is at least one shape in the active publication.




```vb
Sub SetNewTextureForAllShapes() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 shp.Fill.PresetTextured PresetTexture:=msoTextureOak 
 Next shp 
End Sub
```

This example adds a shadow to the first shape in the active publication. This example assumes there is at least one shape in the active publication.




```vb
Sub SetShadowForFirstShape() 
 ActiveDocument.Pages(1).Shapes(1).Shadow.Type = msoShadow6 
End Sub
```

This example displays a count of all shapes on the first page of the active publication. This example assumes there is at least one shape in the active publication.




```vb
Sub CountShapesOnFirstPage() 
 MsgBox "You have " &; ActiveDocument.Pages(1) _ 
 .Shapes.Count &; " shapes on the first page." 
End Sub
```


