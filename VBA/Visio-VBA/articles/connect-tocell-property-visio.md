---
title: Connect.ToCell Property (Visio)
keywords: vis_sdr.chm10314545
f1_keywords:
- vis_sdr.chm10314545
ms.prod: visio
api_name:
- Visio.Connect.ToCell
ms.assetid: 2210e427-132d-d713-02bf-0fd19ce225b7
ms.date: 06/08/2017
---


# Connect.ToCell Property (Visio)

Gets the cell to which a connection is made. Read-only.


## Syntax

 _expression_ . **ToCell**

 _expression_ A variable that represents a **Connect** object.


### Return Value

Cell


## Remarks

A connection is defined by a reference in a cell in the shape from which the connection originates to a cell in the shape to which the connection is made. The  **ToCell** property returns the **Cell** object to which the connection is made.

Following is a list of possible connections and their related  **ToCell** property values.

 **From the begin or end cell of a 1-D shape to...**




-  **A connection point cell:** The **ToCell** property returns the Connection.X _i_ cell.
    
-  **A cell of a guide or guide point:** When the begin or end cell of a 1-D shape is glued to a cell of a guide or guide point, two connections are created—one from the endpoint's X cell to the guide's Angle cell, and the other from the endpoint's Y cell to the guide's Angle cell. The **ToCell** property of both **Connect** objects returns the Angle cell.
    
-  **The pin of a 2-D shape (creates dynamic glue):** The **ToCell** property returns the PinX cell.
    
-  **Any cell of a vertex row in a Geometry section:** A new connection point is created and the **ToCell** property returns the Connections.X _i_ cell.
    
-  **The begin or end cell of a 1-D shape:** A new connection point is created and the **ToCell** property returns the Connections.X _i_ cell.
    
-  **The edge (a cell in the Alignment section) of a 2-D shape:** A new connection point is created and the **ToCell** property returns the Connections.X _i_ cell.
    


 **From the edge (a cell in the Alignment section) of a 2-D shape to acell of a guide or guide point:** The **ToCell** property returns the Angle cell.

 **From an outward or inward/outward connection point cell of a 1-D shape to An inward or inward/outward connection point cell that is not a cell of a guide or guide point:** When these cells are glued, two connections are created—one from the BeginX cell of the 1-D shape to the Connections.X _i_ cell, and the other from the EndX cell of the 1-D shape to the Connections.Y _i_ cell. The **ToCell** property returns Connections.X _i_ for the first **Connect** object and Connections.Y _i_ for the other.

 **From an outward or inward/outward connection point cell of a 2-D shape that is not a guide or guide point to an inward or inward/outward type connection point cell that is not a cell of a guide or guide point:** If the outward connection point is directionless, the **ToCell** property returns the Connections.X _i_ cell. If the outward connection point has a direction, two connection points are created. The **ToCell** property returns the Connections.X _i_ cell for both **Connect** objects.

 **From a control handle to...**




-  **A connection point cell:** The **ToCell** property returns the Connections.Xi cell.
    
-  **A cell of a guide or guide point:** When a control point is glued to a cell of a guide or guide point, two connections are created—one from the control point's X cell to the guide's PinX and the other from the control point's Y cell to the guide's PinY cell. The **ToCell** property of the first **Connect** object returns the guide's PinX cell and, for the second **Connect** object, the guide's PinY cell.
    
-  **Any cell of a vertex row in a Geometry section:** A new connection point is created and the **ToCell** property returns the Connections.Xi cell.
    
-  **The begin or end cell of a 1-D shape:** A new connection point is created and the **ToCell** property returns the Connections.Xi cell.
    
-  **The edge (a cell in the Alignment section) of a 2-D shape:** A new connection point is created and the **ToCell** property returns the Connections.X _i_ cell.
    



## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to extract connection information from a Microsoft Visio drawing. The example displays the connection information in the Immediate window.



This example assumes there is an active document that contains at least two connected shapes.




```vb
 
Public Sub ToCell_Example() 
 
 Dim vso1DShape As Visio.Shape 
 Dim vso2DShape1 As Visio.Shape 
 Dim vso2DShape2 As Visio.Shape 
 Dim vsoCellGlueFromBegin As Visio.Cell 
 Dim vsoCellGlueFromEnd As Visio.Cell 
 Dim vsoCellGlueToObject As Visio.Cell 
 Dim vsoCellGlueToObject2 As Visio.Cell 
 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoShape As Visio.Shape 
 Dim vsoConnects As Visio.Connects 
 Dim vsoConnect As Visio.Connect 
 Dim vsoConnectToCell As Visio.Cell 
 Dim intCurrentShapeID As Integer 
 Dim intCounter As Integer 
 
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
 
 Set vsoShapes = ActivePage.Shapes 
 
 'For each shape on the page, get its connections. 
 For intCurrentShapeID = 1 To vsoShapes.Count 
 
 Set vsoShape = vsoShapes(intCurrentShapeID) 
 Set vsoConnects = vsoShape.Connects 
 
 'For each connection, get the cell it connects to. 
 For intCounter = 1 To vsoConnects.Count 
 
 Set vsoConnect = vsoConnects(intCounter) 
 Set vsoConnectToCell = vsoConnect.ToCell 
 
 'Print connect information in the Immediate window. 
 Debug.Print " To "; vsoConnectToCell.Name 
 
 Next intCounter 
 
 Next intCurrentShapeID 
 
End Sub
```


