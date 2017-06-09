---
title: ConnectorFormat.EndConnectedShape Property (Excel)
keywords: vbaxl10.chm646081
f1_keywords:
- vbaxl10.chm646081
ms.prod: excel
api_name:
- Excel.ConnectorFormat.EndConnectedShape
ms.assetid: e13d9b94-aa51-5895-8ad4-c40ba7397331
ms.date: 06/08/2017
---


# ConnectorFormat.EndConnectedShape Property (Excel)

Returns a  **[Shape](shape-object-excel.md)** object that represents the shape that the end of the specified connector is attached to. Read-only.


## Syntax

 _expression_ . **EndConnectedShape**

 _expression_ A variable that represents a **ConnectorFormat** object.


## Remarks

If the end of the specified connector isn't attached to a shape, this property generates an error.


## Example

This example assumes that  `myDocument` already contains two shapes attached by a connector named "Conn1To2." The code adds a rectangle and a connector to `myDocument`. The end of the new connector will be attached to the same connection site as the end of the connector named "Conn1To2," and the beginning of the new connector will be attached to connection site one on the new rectangle.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 Set r3 = .AddShape(msoShapeRectangle, _ 
 100, 420, 200, 100) 
 With .Item("Conn1To2").ConnectorFormat 
 endConnSite1 = .EndConnectionSite 
 Set endConnShape1 = .EndConnectedShape 
 End With 
 With .AddConnector(msoConnectorCurve, _ 
 0, 0, 10, 10).ConnectorFormat 
 .BeginConnect r3, 1 
 .EndConnect endConnShape1, endConnSite1 
 End With 
End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-excel.md)

