---
title: ConnectorFormat.BeginConnectionSite Property (Excel)
keywords: vbaxl10.chm646079
f1_keywords:
- vbaxl10.chm646079
ms.prod: excel
api_name:
- Excel.ConnectorFormat.BeginConnectionSite
ms.assetid: 606f6e75-3375-da45-b177-63318ef5f594
ms.date: 06/08/2017
---


# ConnectorFormat.BeginConnectionSite Property (Excel)

Returns an integer that specifies the connection site that the beginning of a connector is connected to. Read-only  **Long** .


## Syntax

 _expression_ . **BeginConnectionSite**

 _expression_ A variable that represents a **ConnectorFormat** object.


## Remarks

If the beginning of the specified connector isn't attached to a shape, this property generates an error.


## Example

This example assumes that  `myDocument` already contains two shapes attached by a connector named "Conn1To2." The code adds a rectangle and a connector to `myDocument`. The beginning of the new connector will be attached to the same connection site as the beginning of the connector named "Conn1To2," and the end of the new connector will be attached to connection site one on the new rectangle.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 Set r3 = .AddShape(msoShapeRectangle, 450, 190, 200, 100) 
 .AddConnector(msoConnectorCurve, 0, 0, 10, 10).Name = _ 
 "Conn1To3" 
 With .Item("Conn1To2").ConnectorFormat 
 beginConnSite1 = .BeginConnectionSite 
 Set beginConnShape1 = .BeginConnectedShape 
 End With 
 With .Item("Conn1To3").ConnectorFormat 
 .BeginConnect beginConnShape1, beginConnSite1 
 .EndConnect r3, 1 
 End With 
End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-excel.md)

