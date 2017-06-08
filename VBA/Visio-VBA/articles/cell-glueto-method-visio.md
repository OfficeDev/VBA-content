---
title: Cell.GlueTo Method (Visio)
keywords: vis_sdr.chm10116335
f1_keywords:
- vis_sdr.chm10116335
ms.prod: visio
api_name:
- Visio.Cell.GlueTo
ms.assetid: dc88ecf1-d7c2-994e-8b49-e4bfddef4472
ms.date: 06/08/2017
---


# Cell.GlueTo Method (Visio)

Glues one shape to another, from a cell in the first shape to a cell in the second shape.


## Syntax

 _expression_ . **GlueTo**( **_CellObject_** )

 _expression_ A variable that represents a **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CellObject_|Required| **[IVCELL]**|An expression that returns a  **Cell** object that represents the part of the shape to glue to.|

### Return Value

Nothing


## Remarks

Following is a list of possible connections.

Following is a list of possible connections.


### From the begin or end cell of a 1-D shape to...




- A connection point cell.
    
-  **A cell of a guide or guide point:** When the begin or end cell of a 1-D shape is glued to a cell of a guide or guide point, two connections are created?one from the BeginX or EndX cell to the guide's Angle cell, and one from the BeginY or EndY cell to the guide's Angle cell.
    
-  **The pin of a 2-D shape (creates dynamic glue):** The shape being glued from must be routable (ObjType includes **visLOFlagsRoutable** ) or have a dynamic glue type (GlueType includes **visGlueTypeWalking** ), and does not prohibit dynamic glue (GlueType does not include **visGlueTypeNoWalking** ). Gluing to PinX creates dynamic glue with a horizontal walking preference and gluing to PinY creates dynamic glue with a vertical walking preference.
    
-  **Any cell of a vertex row in a Geometry section:** A connection point is created that is glued to. Either the begin or end cell can be designated as the cell to glue from. The **GlueTo** method establishes new formulas in both the X and Y cells of the connection row.
    
-  **The begin or end cell of a 1-D shape:** A connection point is created that is glued to. Either the begin or end cell can be designated as the cell to glue from. The **GlueTo** method establishes new formulas in both the X and Y cells of the connection row.
    
-  **The edge (a cell in the Alignment section) of a 2-D shape:** A connection point is created that is glued to. Either the begin or end cell can be designated as the cell to glue from. The **GlueTo** method establishes new formulas in both the X and Y cells of the connection row.
    
 **From the edge (a cell in the Alignment section) of a 2-D shape to a cell of a guide or guide point.**

 **From an outward or inward/outward connection point cell of a 1-D shape to an inward or inward/outward connection point cell that is not a cell of a guide or guide point:** When these cells are glued, two connections are created?one from the BeginX cell of the 1-D shape to the Connections.Xi cell, and the other from the EndX cell of the 1-D shape to the Connections.Y _i_ cell.

 **From an outward or inward/outward connection point cell of a 2-D shape to an inward or inward/outward type connection point cell that is not a cell of a guide or guide point:** If the outward connection point has a direction, two connection points are created?one from the Angle cell to the Connections.X _i_ cell and the other from the PinX cell to the Connections.Y _i_ cell.


### From a control point cell to...




- A connection point cell.
    
-  **A cell of a guide or guide point:** When a control point is glued to a cell of a guide or guide point other than a connection point cell, two connections are created?one to the guide's PinX and one to the guide's PinY.
    
-  **Any cell of a vertex row in a Geometry section:** A connection point is created that is glued to. Any cell in the control point row can be designated as the cell to glue from. The **GlueTo** method establishes new formulas in both the X and Y cells of the connection row.
    
-  **The begin or end cell of a 1-D shape that isn't a guide or guide point:** A connection point is created that is glued to. Any cell in the control point row can be designated as the cell to glue from. The **GlueTo** method establishes new formulas in both the X and Y cells of the connection row.
    
-  **The edge (a cell in the Alignment section) of a 2-D shape:** A connection point is created that is glued to. Any cell in the control point row can be designated as the cell to glue from. The **GlueTo** method establishes new formulas in both the X and Y cells of the connection row.
    
For details about connection point type and direction, see the Connection Points section.


## Example

The following macro shows how to use the  **GlueTo** method to glue shapes together.


```vb
 
Public Sub GlueTo_Example()  
 
    Dim vso1DShape As Visio.Shape  
    Dim vso2DShape1 As Visio.Shape  
    Dim vso2DShape2 As Visio.Shape  
    Dim vsoCellGlueFromBegin As Visio.Cell  
    Dim vsoCellGlueFromEnd As Visio.Cell  
    Dim vsoCellGlueToObject As Visio.Cell  
    Dim vsoCellGlueToObject2 As Visio.Cell  
 
    'Draw a line.  
    Set vso1DShape = ActivePage.DrawLine(3, 5, 5, 3)  
 
    'Draw the lower rectangle.  
    Set vso2DShape1 = ActivePage.DrawRectangle(1, 1, 4, 2)  
 
    'Draw the upper rectangle.  
    Set vso2DShape2 = ActivePage.DrawRectangle(5, 5, 8, 6)  
 
    'Get the Cell objects needed to make the connections.  
    Set vsoCellGlueFromBegin = vso1DShape.Cells("BeginX")  
    Set vsoCellGlueFromEnd = vso1DShape.Cells("EndX")  
    Set vsoCellGlueToObject = vso2DShape1.Cells("Geometry1.X3")  
    Set vsoCellGlueToObject2 = vso2DShape2.Cells("Geometry1.X1")  
 
    'Use the GlueTo method to glue the begin point of the 1-D shape  
    'to the top right vertex (Geometry1.X3) of the lower 2-D shape.  
    vsoCellGlueFromBegin.GlueTo vsoCellGlueToObject  
 
    'Use the GlueTo method to glue the endpoint of the 1-D shape  
    'to the bottom left vertex (Geometry1.X1) of the upper 2-D shape.  
    vsoCellGlueFromEnd.GlueTo vsoCellGlueToObject2  
 
    'You can also use the GlueTo method to glue  
    'by referencing a connection point cell.  
    Set vso1DShape = ActivePage.DrawLine(3, 5, 5, 3)  
    Set vsoCellGlueFromEnd = vso1DShape.Cells("EndX")  
    Set vsoCellGlueToObject = vso2DShape1.Cells("Connections.X1")  
    vsoCellGlueFromEnd.GlueTo vsoCellGlueToObject  
 
End Sub
```


