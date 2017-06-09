---
title: ShapeRange.RerouteConnections Method (Publisher)
keywords: vbapb10.chm2293796
f1_keywords:
- vbapb10.chm2293796
ms.prod: publisher
api_name:
- Publisher.ShapeRange.RerouteConnections
ms.assetid: ae2a64ea-1b7a-4ff6-304c-680dd96fd386
ms.date: 06/08/2017
---


# ShapeRange.RerouteConnections Method (Publisher)

Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the  **RerouteConnections** method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.


## Syntax

 _expression_. **RerouteConnections**

 _expression_A variable that represents a  **ShapeRange** object.


### Return Value

Nothing


## Remarks

This method reroutes all connectors attached to the specified shape; if the specified shape is a connector, it is rerouted.


## Example

This example adds two rectangles to the first page in the active publication and connects them with a curved connector. Note that the  **RerouteConnections** method overrides the values you supply for the **_ConnectionSite_** arguments used with the **BeginConnect**and  **EndConnect** methods.


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


