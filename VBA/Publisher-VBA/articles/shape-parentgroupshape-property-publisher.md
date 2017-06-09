---
title: Shape.ParentGroupShape Property (Publisher)
keywords: vbapb10.chm2228338
f1_keywords:
- vbapb10.chm2228338
ms.prod: publisher
api_name:
- Publisher.Shape.ParentGroupShape
ms.assetid: ced4c348-4ef5-c703-fdea-65c33d37b4c0
ms.date: 06/08/2017
---


# Shape.ParentGroupShape Property (Publisher)

Returns a  **[Shape](shape-object-publisher.md)** object that represents the common parent shape of a child shape or a range of child shapes.


## Syntax

 _expression_. **ParentGroupShape**

 _expression_A variable that represents a  **Shape** object.


### Return Value

Shape


## Example

This example creates two shapes in the active document and groups those shapes. Then using one shape in the group, it accesses the parent group and fills all shapes in the parent group with the same fill pattern. This example assumes that the active document does not currently contain any shapes. If it does, an error may occur.


```vb
Sub ParentGroupShape() 
 Dim shpGroup As Shape 
 
 With ActiveDocument.Pages(1).Shapes 
 .AddShape Type:=msoShapeOval, Left:=72, _ 
 Top:=72, Width:=100, Height:=100 
 .AddShape Type:=msoShapeHeart, Left:=110, _ 
 Top:=120, Width:=100, Height:=100 
 .Range(Array(1, 2)).Group 
 End With 
 
 Set shpGroup = ActiveDocument.Pages(1).Shapes(1) _ 
 .GroupItems(1).ParentGroupShape 
 shpGroup.Fill.Patterned Pattern:=msoPattern25Percent 
 
End Sub
```


