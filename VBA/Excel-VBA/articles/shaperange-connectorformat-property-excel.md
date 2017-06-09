---
title: ShapeRange.ConnectorFormat Property (Excel)
keywords: vbaxl10.chm640102
f1_keywords:
- vbaxl10.chm640102
ms.prod: excel
api_name:
- Excel.ShapeRange.ConnectorFormat
ms.assetid: cc2c9559-a7f5-8e32-1976-c81e400fb9dd
ms.date: 06/08/2017
---


# ShapeRange.ConnectorFormat Property (Excel)

Returns a  **[ConnectorFormat](connectorformat-object-excel.md)** object that contains connector formatting properties. Applies to a **[ShapeRange](shaperange-object-excel.md)** objects that represent connectors. Read-only.


## Syntax

 _expression_ . **ConnectorFormat**

 _expression_ An expression that returns a **ShapeRange** object.


## Example

This example adds two rectangles to  `myDocument`, attaches them with a connector, automatically reroutes the connector along the shortest path, and then detaches the connector from the rectangles.


```vb
Set myDocument = Worksheets(1) 
Set s = myDocument.Shapes 
Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) 
Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) 
Set c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0) 
with c.ConnectorFormat 
 .BeginConnect firstRect, 1 
 .EndConnect secondRect, 1 
 c.RerouteConnections 
 .BeginDisconnect 
 .EndDisconnect 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

