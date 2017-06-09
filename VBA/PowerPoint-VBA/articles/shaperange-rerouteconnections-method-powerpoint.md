---
title: ShapeRange.RerouteConnections Method (PowerPoint)
keywords: vbapp10.chm548009
f1_keywords:
- vbapp10.chm548009
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.RerouteConnections
ms.assetid: 61db5f5d-74cd-1b9d-1b37-9d33e320cca8
ms.date: 06/08/2017
---


# ShapeRange.RerouteConnections Method (PowerPoint)

Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the  **RerouteConnections** method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.


## Syntax

 _expression_. **RerouteConnections**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

This method reroutes all connectors attached to the specified shape; if the specified shape is a connector, it is rerouted.

If this method is applied to a connector, only that connector will be rerouted. If this method is applied to a connected shape, all connectors to that shape will be rerouted.


## Example

This example adds two rectangles to  `myDocument`, connects them with a curved connector, and then reroutes the connector so that it takes the shortest possible path between the two rectangles. Note that the  **RerouteConnections** method adjusts the size and position of the connector and determines which connecting sites it attaches to, so the values you initially specify for the ConnectionSite arguments used with the **BeginConnect** and **EndConnect** methods are irrelevant.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set s = myDocument.Shapes

Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
Set newConnector = s _
    .AddConnector(msoConnectorCurve, 0, 0, 100, 100)

With newConnector.ConnectorFormat
    .BeginConnect firstRect, 1
    .EndConnect secondRect, 1
End With

newConnector.RerouteConnections
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

