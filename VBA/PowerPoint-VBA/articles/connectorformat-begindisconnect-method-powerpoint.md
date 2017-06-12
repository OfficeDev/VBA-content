---
title: ConnectorFormat.BeginDisconnect Method (PowerPoint)
keywords: vbapp10.chm555003
f1_keywords:
- vbapp10.chm555003
ms.prod: powerpoint
api_name:
- PowerPoint.ConnectorFormat.BeginDisconnect
ms.assetid: 8f556e09-b874-73b8-902a-2446ddedd0f4
ms.date: 06/08/2017
---


# ConnectorFormat.BeginDisconnect Method (PowerPoint)

Detaches the beginning of the specified connector from the shape it is attached to. 


## Syntax

 _expression_. **BeginDisconnect**

 _expression_ A variable that represents a **ConnectorFormat** object.


## Remarks

This method doesn't alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected. Use the  **[EndDisconnect](connectorformat-enddisconnect-method-powerpoint.md)** method to detach the end of the connector from a shape.


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

