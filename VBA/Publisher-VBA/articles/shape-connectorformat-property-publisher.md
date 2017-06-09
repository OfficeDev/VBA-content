---
title: Shape.ConnectorFormat Property (Publisher)
keywords: vbapb10.chm2228278
f1_keywords:
- vbapb10.chm2228278
ms.prod: publisher
api_name:
- Publisher.Shape.ConnectorFormat
ms.assetid: 280c424c-530c-55ab-da4f-65b858ee3dd8
ms.date: 06/08/2017
---


# Shape.ConnectorFormat Property (Publisher)

Returns a  **[ConnectorFormat](connectorformat-object-publisher.md)** object that contains connector formatting properties. Applies to  **Shape** or **ShapeRange** objects that represent connectors. Read-only.


## Syntax

 _expression_. **ConnectorFormat**

 _expression_A variable that represents a  **Shape** object.


## Example

This example adds two rectangles to the first page in the active publication and connects them with a curved connector.


```vb
Dim shpRect1 As Shape 
Dim shpRect2 As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 
 ' Add two new rectangles. 
 Set shpRect1 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=100, Top:=50, Width:=200, Height:=100) 
 Set shpRect2 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=300, Width:=200, Height:=100) 
 
 ' Add a new curved connector. 
 With .AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=100, EndY:=100) _ 
 .ConnectorFormat 
 
 ' Connect the new connector to the two rectangles. 
 .BeginConnect ConnectedShape:=shpRect1, ConnectionSite:=1 
 .EndConnect ConnectedShape:=shpRect2, ConnectionSite:=1 
 
 ' Reroute the connector to create the shortest path. 
 .Parent.RerouteConnections 
 End With 
 
End With 

```


