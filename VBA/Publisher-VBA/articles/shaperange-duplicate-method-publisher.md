---
title: ShapeRange.Duplicate Method (Publisher)
keywords: vbapb10.chm2293780
f1_keywords:
- vbapb10.chm2293780
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Duplicate
ms.assetid: e940e551-4307-aa33-5713-80f77fade8af
ms.date: 06/08/2017
---


# ShapeRange.Duplicate Method (Publisher)

Creates a duplicate of the specified  **[Shape](shape-object-publisher.md)** or **[ShapeRange](shaperange-object-publisher.md)** object, adds the new shape or range of shapes to the **Shapes** collection immediately after the shape or range of shapes specified originally, and then returns the new **Shape** or **ShapeRange** object.


## Syntax

 _expression_. **Duplicate**

 _expression_A variable that represents a  **ShapeRange** object.


### Return Value

ShapeRange


## Example

This example adds a new, blank page at the end of the active publication, adds a diamond shape to the new page, duplicates the diamond, and then sets properties for the duplicate. The first diamond will have the default fill color for the active color scheme; the second diamond will be offset from the first one and will have the first accent color for the active color scheme.


```vb
Dim pgTemp As Page 
Dim shpTemp As Shape 
 
Set pgTemp = ActiveDocument.Pages.Add(Count:=1, After:=1) 
Set shpTemp = pgTemp.Shapes _ 
 .AddShape(Type:=msoShapeDiamond, _ 
 Left:=10, Top:=10, Width:=250, Height:=350) 
 
With shpTemp.Duplicate 
 .Left = 150 
 .Fill.ForeColor.SchemeColor = pbSchemeColorAccent1 
End With
```


