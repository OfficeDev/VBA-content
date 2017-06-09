---
title: ConnectorFormat.BeginDisconnect Method (Excel)
keywords: vbaxl10.chm646074
f1_keywords:
- vbaxl10.chm646074
ms.prod: excel
api_name:
- Excel.ConnectorFormat.BeginDisconnect
ms.assetid: 1edd106a-9f02-3916-401c-1b026e40d75a
ms.date: 06/08/2017
---


# ConnectorFormat.BeginDisconnect Method (Excel)

Detaches the beginning of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected. Use the  **[EndDisconnect](connectorformat-enddisconnect-method-excel.md)** method to detach the end of the connector from a shape.


## Syntax

 _expression_ . **BeginDisconnect**

 _expression_ A variable that represents a **ConnectorFormat** object.


## Example

This example adds two rectangles to  `myDocument`, attaches them with a connector, automatically reroutes the connector along the shortest path, and then detaches the connector from the rectangles.


```vb
Set myDocument = Worksheets(1) 
Set s = myDocument.Shapes 
Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) 
Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) 
Set c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0) 
With c.ConnectorFormat 
 .BeginConnect firstRect, 1 
 .EndConnect secondRect, 1 
 c.RerouteConnections 
 .BeginDisconnect 
 .EndDisconnect 
End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-excel.md)

