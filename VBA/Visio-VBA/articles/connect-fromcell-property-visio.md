---
title: Connect.FromCell Property (Visio)
keywords: vis_sdr.chm10313575
f1_keywords:
- vis_sdr.chm10313575
ms.prod: visio
api_name:
- Visio.Connect.FromCell
ms.assetid: d605d25a-40c2-7e7c-c8c2-bbc31c00f47b
ms.date: 06/08/2017
---


# Connect.FromCell Property (Visio)

Returns the cell from which a connection originates. Read-only.


## Syntax

 _expression_ . **FromCell**

 _expression_ A variable that represents a **Connect** object.


### Return Value

Cell


## Remarks

A connection is defined by a reference in a cell in the shape from which the connection originates to a cell in the shape to which the connection is made. The  **FromCell** property returns the **Cell** object for the cell from which the connection originates.

Following is a list of possible connections and the values of their related  **FromCell** properties.

A connection is defined by a reference in a cell in the shape from which the connection originates to a cell in the shape to which the connection is made. The  **FromCell** property returns the **Cell** object for the cell from which the connection originates.

Following is a list of possible connections and the values of their related  **FromCell** properties.


### From the begin or end cell of a 1-D shape to...




-  **A connection point cell:** The **FromCell** property returns either the BeginX or EndX cell, depending on which endpoint was glued.
    
-  **A cell of a guide or guide point:** When the begin or end cell of a 1-D shape is glued to a cell of a guide or guide point, two connections are created—one from the endpoint's X cell to the guide's Angle cell, and one from the endpoint's Y cell to the guide's Angle cell. The **FromCell** property of one **Connect** object returns the BeginX or EndX cell and, the **FromCell** property of the other **Connect** object returns the BeginY or EndY cell, depending on which endpoint is glued.
    
-  **The pin of a 2-D shape (creates dynamic glue):** The shape from which the glue originates must be routable or have a dynamic glue type. The **FromCell** property returns either the BeginX or EndX cell, depending on which endpoint was glued.
    
-  **Any cell of a vertex row in a Geometry section:** The **FromCell** property returns either the BeginX or EndX cell, depending on which endpoint was glued.
    
-  **The begin or end cell of a 1-D shape:** The **FromCell** property returns either the BeginX or EndX cell, depending on which endpoint was glued.
    
-  **The edge (a cell in the Alignment section) of a 2-D shape:** The **FromCell** property returns either the BeginX or EndX cell, depending on which endpoint was glued.
    

### From the edge (a cell in the Alignment section) of a 2-D shape to a cell of a guide or guide point:

The  **FromCell** property returns the Alignment cell that is glued to the guide.


### From an outward or inward/outward connection point cell of a 1-D shape to an inward or inward/outward connection point cell that is not a cell of a guide or guide point:

 When these cells are glued, two connections are created—one from the BeginX cell of the 1-D shape to the Connections.X _i_ cell, and the other from the EndX cell of the 1-D shape to the Connections.Y _i_ cell. The **FromCell** property returns BeginX for one **Connect** object, and EndX for the other.


### From an outward or inward/outward connection point cell of a 2-D shape to an inward or inward/outward type connection point cell that is not a cell of a guide or guide point:

 If the outward connection point is directionless, the **FromCell** property returns the PinX cell. If the outward connection point has a direction, then two connection points are created. The **FromCell** property returns the Angle cell for one **Connect** object and the PinX cell for the other.


### From a control point cell to...




-  **A connection point cell:** The **FromCell** property returns the Controls.X _i_ cell.
    
-  **A cell of a guide or guide point:** When a control point is glued to a cell of a guide or guide point other than a connection point cell, two connections are created. The **FromCell** property of one **Connect** object returns Controls.X _i_ and the second **Connect** object returns Controls.Y _i_ .
    
-  **Any cell of a vertex row in a Geometry section:** The **FromCell** property returns the Controls.X _i_ cell.
    
-  **The begin or end cell of a 1-D shape that isn't a guide or guide point:** The **FromCell** property returns the Controls.X _i_ cell.
    
-  **The edge (a cell in the Alignment section) of a 2-D shape:** The **FromCell** property returns the Controls.X _i_ cell.
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to extract connection information from a Microsoft Visio drawing. The example displays the connection information in the Immediate window.



This example assumes there is an active document that contains at least two connected shapes.




```vb
 
Public Sub FromCell_Example() 
  
    Dim vsoShapes As Visio.Shapes  
    Dim vsoShape As Visio.Shape  
    Dim vsoConnectCell As Visio.Cell  
    Dim vsoConnects As Visio.Connects  
    Dim vsoConnect As Visio.Connect  
    Dim intCurrentShapeID As Integer 
    Dim intCounter As Integer 
    Set vsoShapes = ActivePage.Shapes 
  
    'For each shape on the page, get all its connections. 
    For intCurrentShapeIndex = 1 To vsoShapes.Count  
        Set vsoShape = vsoShapes(intCurrentShapeIndex)  
        Set vsoConnects = vsoShape.Connects  
  
        'For each connection, get the cell the connection  
        'originates from, and print its name in the Immediate window. 
        For intCounter = 1 To vsoConnects.Count  
            Set vsoConnect = vsoConnects(intCounter)  
            Set vsoConnectCell = vsoConnect.FromCell  
            Debug.Print "From " &; vsoConnectCell.Name  
        Next intCounter  
 
    Next intCurrentShapeIndex  
 
End Sub
```


