---
title: ConnectorFormat.BeginConnect Method (Publisher)
keywords: vbapb10.chm3211280
f1_keywords:
- vbapb10.chm3211280
ms.prod: publisher
api_name:
- Publisher.ConnectorFormat.BeginConnect
ms.assetid: d38f6ac7-f09b-b171-a6b8-d52427f45d78
ms.date: 06/08/2017
---


# ConnectorFormat.BeginConnect Method (Publisher)

Attaches the beginning of the specified connector to a specified shape.


## Syntax

 _expression_. **BeginConnect**( **_ConnectedShape_**,  **_ConnectionSite_**)

 _expression_A variable that represents a  **ConnectorFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ConnectedShape|Required| **Shape**|The shape to which Microsoft Publisher attaches the beginning of the connector. The specified  **Shape** object must be in the same **Shapes** collection as the connector.|
|ConnectionSite|Required| **Long**|A connection site on the shape specified by ConnectedShape. Must be an integer between 1 and the integer returned by the  **[ConnectionSiteCount](shape-connectionsitecount-property-publisher.md)** property of the specified shape. Connection sites are numbered starting from the top of the specified shape and moving counterclockwise around the shape. If you want the connector to automatically find the shortest path between the two shapes it connects, specify any valid integer for this argument and then use the **[RerouteConnections](shape-rerouteconnections-method-publisher.md)** method after the connector is attached to shapes at both ends.|

## Remarks

If there's already a connection between the beginning of the connector and another shape, that connection is broken. If the beginning of the connector isn't already positioned at the specified connecting site, this method moves the beginning of the connector to the connecting site and adjusts the size and position of the connector.

When you attach a connector to an object, the size and position of the connector are automatically adjusted if necessary.

Use the  **[EndConnect](connectorformat-endconnect-method-publisher.md)** method to attach the end of the connector to a shape.


## Example

This example adds two rectangles to the first page in the active publication and connects them with a curved connector. Note that the  **RerouteConnections** method overrides the values you supply for the **_ConnectionSite_** arguments used with the **BeginConnect** and **EndConnect** methods.


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


