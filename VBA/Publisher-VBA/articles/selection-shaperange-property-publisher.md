---
title: Selection.ShapeRange Property (Publisher)
keywords: vbapb10.chm851972
f1_keywords:
- vbapb10.chm851972
ms.prod: publisher
api_name:
- Publisher.Selection.ShapeRange
ms.assetid: d95cce6d-e3a2-09b9-a6d5-749e0476544c
ms.date: 06/08/2017
---


# Selection.ShapeRange Property (Publisher)

Returns a  **[ShapeRange](shaperange-object-publisher.md)** collection that represents all the **Shape** objects in the specified range or selection. The shape range can contain drawings, shapes, pictures, OLE objects, ActiveX controls, text objects, and callouts.


## Syntax

 _expression_. **ShapeRange**

 _expression_A variable that represents a  **Selection** object.


### Return Value

ShapeRange


## Example

The following example sets the fill pattern for all the shapes in the selection. This example assumes one or more shapes are selected in the active publication.


```vb
Sub ChangeFillForShapeRange() 
 Selection.ShapeRange.Fill.Patterned Pattern:=msoPattern20Percent 
End Sub
```

The following example applies shadow and fill formatting to all the shapes in the selection. This example assumes one or more shapes are selected in the active publication.




```vb
Sub SetShadowForSelectedShapes() 
 With Selection.ShapeRange 
 .Shadow.Type = msoShadow6 
 .Fill.Patterned Pattern:=msoPatternDottedDiamond 
 End With 
End Sub
```


