---
title: ConnectorFormat Object (Publisher)
keywords: vbapb10.chm3276799
f1_keywords:
- vbapb10.chm3276799
ms.prod: publisher
api_name:
- Publisher.ConnectorFormat
ms.assetid: 9b541d54-b1b9-c023-c9c4-08ff6b811eb9
ms.date: 06/08/2017
---


# ConnectorFormat Object (Publisher)

Contains properties and methods that apply to connectors. A connector is a line that attaches two other shapes at points called connection sites. If you rearrange shapes that are connected, the geometry of the connector will be automatically adjusted so that the shapes remain connected.
 


## Example

Use the  **ConnectorFormat** property of the **[Shape](shape-object-publisher.md)** object or the **[ShapeRange](shaperange-object-publisher.md)** collection to return a **ConnectorFormat** object. Use the **[BeginConnect](connectorformat-beginconnect-method-publisher.md)** and **[EndConnect](connectorformat-endconnect-method-publisher.md)** methods of the **ConnectorFormat** object to attach the ends of the connector to other shapes in the publication. Use the **[RerouteConnections](shape-rerouteconnections-method-publisher.md)** method of the **Shape** object and **ShapeRange** collection to automatically find the shortest path between the two shapes connected by the connector. Use the **[Connector](shape-connector-property-publisher.md)** property to see whether a shape is a connector.
 

 

 

 
Note that you assign a size and a position when you add a connector to the  **Shapes** collection, but the size and position are automatically adjusted when you attach the beginning and end of the connector to other shapes in the collection. Therefore, if you intend to attach a connector to other shapes, the initial size and position you specify are irrelevant. Likewise, you specify which connection sites on a shape to attach the connector to when you attach the connector, but using the **RerouteConnections** method after the connector is attached may change which connection sites the connector attaches to, making your original choice of connection sites irrelevant.
 

 

 

 
The following example adds two rectangles to the active publication and connects them with a curved connector.
 

 



```
Dim shpAll As Shapes 
Dim firstRect As Shape 
Dim secondRect As Shape 
 
Set shpAll = ActiveDocument.Pages(1).Shapes 
Set firstRect = shpAll.AddShape(Type:=msoShapeRectangle, _ 
 Left:=100, Top:=50, Width:=200, Height:=100) 
Set secondRect = shpAll.AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=300, Width:=200, Height:=100) 

```




```
With shpAll.AddConnector(Type:=msoConnectorCurve, BeginX:=0, _ 
 BeginY:=0, EndX:=0, EndY:=0).ConnectorFormat 
 .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1 
 .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1 
 .Parent.RerouteConnections 
End With
```


## Methods



|**Name**|
|:-----|
|[BeginConnect](connectorformat-beginconnect-method-publisher.md)|
|[BeginDisconnect](connectorformat-begindisconnect-method-publisher.md)|
|[EndConnect](connectorformat-endconnect-method-publisher.md)|
|[EndDisconnect](connectorformat-enddisconnect-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](connectorformat-application-property-publisher.md)|
|[BeginConnected](connectorformat-beginconnected-property-publisher.md)|
|[BeginConnectedShape](connectorformat-beginconnectedshape-property-publisher.md)|
|[BeginConnectionSite](connectorformat-beginconnectionsite-property-publisher.md)|
|[EndConnected](connectorformat-endconnected-property-publisher.md)|
|[EndConnectedShape](connectorformat-endconnectedshape-property-publisher.md)|
|[EndConnectionSite](connectorformat-endconnectionsite-property-publisher.md)|
|[Parent](connectorformat-parent-property-publisher.md)|
|[Type](connectorformat-type-property-publisher.md)|

