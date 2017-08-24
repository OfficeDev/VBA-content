---
title: ConnectorFormat.BeginDisconnect Method (Publisher)
keywords: vbapb10.chm3211281
f1_keywords:
- vbapb10.chm3211281
ms.prod: publisher
api_name:
- Publisher.ConnectorFormat.BeginDisconnect
ms.assetid: 30d8ffc0-e8a5-6d9e-a098-8c06d5fde3a9
ms.date: 06/08/2017
---


# ConnectorFormat.BeginDisconnect Method (Publisher)

Detaches the beginning of the specified connector from the shape to which it is attached.


## Syntax

 _expression_. **BeginDisconnect**

 _expression_A variable that represents a  **ConnectorFormat** object.


## Remarks

This method doesn't alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected.

Use the  **[EndDisconnect](connectorformat-enddisconnect-method-publisher.md)** method to detach the end of the connector from a shape.


## Example

This example adds two rectangles to the first page in the active publication, attaches them with a connector, automatically reroutes the connector along the shortest path, and then detaches the connector from the rectangles.


```vb
Dim shpRect1 As Shape 
Dim shpRect2 As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 
 ' Add two new rectangles. 
 Set shpRect1 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=100, Top:=50, Width:=200, Height:=100) 
 Set shpRect2 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=300, Width:=200, Height:=100) 
 
 ' Add a new connector. 
 With .AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=0, EndY:=0) _ 
 .ConnectorFormat 
 
 ' Connect the new connector to the two rectangles. 
 .BeginConnect ConnectedShape:=shpRect1, ConnectionSite:=1 
 .EndConnect ConnectedShape:=shpRect2, ConnectionSite:=1 
 
 ' Reroute the connector to create the shortest path. 
 .Parent.RerouteConnections 
 
 ' Disconnect the new connector from the rectangles but 
 ' leave in place. 
 .BeginDisconnect 
 .EndDisconnect 
 End With 
 
End With 

```


