---
title: ConnectorFormat.EndDisconnect Method (PowerPoint)
keywords: vbapp10.chm555005
f1_keywords:
- vbapp10.chm555005
ms.prod: powerpoint
api_name:
- PowerPoint.ConnectorFormat.EndDisconnect
ms.assetid: e26600c4-a384-5c83-96e6-1060f8ce8d21
ms.date: 06/08/2017
---


# ConnectorFormat.EndDisconnect Method (PowerPoint)

Detaches the end of the specified connector from the shape it is attached to. This method doesn't alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected. Use the  **[BeginDisconnect](connectorformat-begindisconnect-method-powerpoint.md)** method to detach the beginning of the connector from a shape.


## Syntax

 _expression_. **EndDisconnect**

 _expression_ A variable that represents a **ConnectorFormat** object.


## Example

This example adds two rectangles to the first slide in the active presentation, attaches them with a connector, automatically reroutes the connector along the shortest path, and then detaches the connector from the rectangles.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set s = myDocument.Shapes

Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)

Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)

With s.AddConnector(msoConnectorCurve, 0, 0, 0, 0).ConnectorFormat

    .BeginConnect firstRect, 1

    .EndConnect secondRect, 1

    .Parent.RerouteConnections

    .BeginDisconnect

    .EndDisconnect

End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-powerpoint.md)

