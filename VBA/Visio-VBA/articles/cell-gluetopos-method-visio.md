---
title: Cell.GlueToPos Method (Visio)
keywords: vis_sdr.chm10116340
f1_keywords:
- vis_sdr.chm10116340
ms.prod: visio
api_name:
- Visio.Cell.GlueToPos
ms.assetid: 9f9e10f2-030f-f7ad-be04-ea2804c20cb4
ms.date: 06/08/2017
---


# Cell.GlueToPos Method (Visio)

Glues one shape to another from a cell in the first shape to an  _x_, _y_ position in the second shape.


## Syntax

 _expression_ . **GlueToPos**( **_SheetObject_** , **_xPercent_** , **_yPercent_** )

 _expression_ A variable that represents a **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SheetObject_|Required| **[IVSHAPE]**|An expression that returns the  **Shape** object to be glued to.|
| _xPercent_|Required| **Double**|The x-coordinate of the position to glue to.|
| _yPercent_|Required| **Double**|The y-coordinate of the position to glue to.|

### Return Value

Nothing


## Remarks

The  **GlueToPos** method creates a new connection point at the location determined by _xPercent_ and _yPercent_, which represent decimal fractions of the specified shape's width and height, respectively, rather than coordinates. For example, the following creates a connection point at the center of  _SheetObject_ and glues the part of the shape that _cellObject_ represents to that point:

 _cellObject_. **GlueToPos**_SheetObject_, 0.5, 0.5

Gluing the X cell of a Controls section row or a BeginX or EndX cell automatically glues the Y cell of the Controls section row or the BeginY or EndY cell, respectively. (The reverse is also true.)


## Example

The following example shows how to use the  **GlueToPos** method to glue shapes together.


```vb
 
Public Sub GlueToPos_Example() 
 
 Dim vso1DShape As Visio.Shape 
 Dim vso2DShape1 As Visio.Shape 
 Dim vso2DShape2 As Visio.Shape 
 Dim vsoCellGlueFromBegin As Visio.Cell 
 Dim vsoCellGlueFromEnd As Visio.Cell 
 
 'Draw a line. 
 Set vso1DShape = ActivePage.DrawLine(3, 5, 5, 3) 
 
 'Draw the lower rectangle. 
 Set vso2DShape1 = ActivePage.DrawRectangle(1, 1, 4, 2) 
 
 'Draw the upper rectangle. 
 Set vso2DShape2 = ActivePage.DrawRectangle(5, 5, 8, 6) 
 
 'Get the Cell objects needed to make the connections. 
 Set vsoCellGlueFromBegin = vso1DShape.Cells("BeginX") 
 Set vsoCellGlueFromEnd = vso1DShape.Cells("EndX") 
 
 'Use the GlueToPos method to glue the begin point of the 1-D shape 
 'to the top center of the lower 2-D shape. 
 vsoCellGlueFromBegin.GlueToPos vso2DShape1, 0.5, 1 
 
 'Use the GlueToPos method to glue the endpoint of the 1-D shape 
 'to the bottom center of the upper 2-D shape. 
 vsoCellGlueFromEnd.GlueToPos vso2DShape2, 0.5, 0 
 
End Sub
```


