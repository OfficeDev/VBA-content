---
title: WizardValue.Parent Property (Publisher)
keywords: vbapb10.chm2097154
f1_keywords:
- vbapb10.chm2097154
ms.prod: publisher
api_name:
- Publisher.WizardValue.Parent
ms.assetid: aa00d170-550c-8584-1e88-eb122763a0d7
ms.date: 06/08/2017
---


# WizardValue.Parent Property (Publisher)

Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.


## Syntax

 _expression_. **Parent**

 _expression_A variable that represents a  **WizardValue** object.


## Example

This example accesses the parent object of the selected shape, and then adds a new shape to it and sets the fill for the new shape.


```vb
Sub ParentObject() 
 Dim shp As Shape 
 Dim pg As Page 
 
 Set pg = Selection.ShapeRange(1).Parent 
 Set shp = pg.Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=72, Top:=72, Width:=72, Height:=72) 
 
 shp.Fill.ForeColor.RGB = RGB(Red:=180, Green:=180, Blue:=180) 
End Sub
```

This example returns the parent object of a text frame, which is the first shape in the active publication, and then fills the shape with a pattern.




```vb
Sub ParentShape() 
 Dim shpParent As Shape 
 Set shpParent = ActiveDocument.Pages(1).Shapes(1).TextFrame.Parent 
 shpParent.Fill.Patterned Pattern:=msoPatternSphere 
End Sub
```


