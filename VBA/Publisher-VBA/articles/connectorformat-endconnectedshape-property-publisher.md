---
title: ConnectorFormat.EndConnectedShape Property (Publisher)
keywords: vbapb10.chm3211524
f1_keywords:
- vbapb10.chm3211524
ms.prod: publisher
api_name:
- Publisher.ConnectorFormat.EndConnectedShape
ms.assetid: 3ee4a031-228a-cac7-8268-29bfc3a888c1
ms.date: 06/08/2017
---


# ConnectorFormat.EndConnectedShape Property (Publisher)

Returns a  **[Shape](shape-object-publisher.md)** object that represents the shape to which the end of the specified connector is attached.


## Syntax

 _expression_. **EndConnectedShape**

 _expression_A variable that represents an  **ConnectorFormat** object.


### Return Value

Shape


## Remarks

If the end of the specified connector isn't attached to a shape, an error occurs.

Use the  **[BeginConnectedShape](connectorformat-beginconnectedshape-property-publisher.md)** property to return the shape attached to the beginning of a connector.


## Example

This example assumes that the first page in the active publication already contains two shapes attached by a connector named Conn1To2. The code adds a rectangle and a connector to the first page. The end of the new connector will be attached to the same connection site as the end of the connector named Conn1To2, and the beginning of the new connector will be attached to connection site one on the new rectangle.


```vb
Dim shpNew As Shape 
Dim intSite As Integer 
Dim shpOld As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 
 ' Add new rectangle. 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=450, Top:=190, Width:=200, Height:=100) 
 
 ' Add new connector. 
 .AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=10, EndY:=10) _ 
 .Name = "Conn1To3" 
 
 ' Get connection site number of old shape, and set 
 ' reference to old shape. 
 With .Item("Conn1To2").ConnectorFormat 
 intSite = .EndConnectionSite 
 Set shpOld = .EndConnectedShape 
 End With 
 
 ' Connect new connector to old shape and new rectangle. 
 With .Item("Conn1To3").ConnectorFormat 
 .EndConnect ConnectedShape:=shpOld, _ 
 ConnectionSite:=intSite 
 .BeginConnect ConnectedShape:=shpNew, _ 
 ConnectionSite:=1 
 End With 
End With 

```


