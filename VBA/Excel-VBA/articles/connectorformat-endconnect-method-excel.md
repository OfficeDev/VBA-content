---
title: ConnectorFormat.EndConnect Method (Excel)
keywords: vbaxl10.chm646075
f1_keywords:
- vbaxl10.chm646075
ms.prod: excel
api_name:
- Excel.ConnectorFormat.EndConnect
ms.assetid: c8cc392c-8a54-99ed-ffdd-e5173792408f
ms.date: 06/08/2017
---


# ConnectorFormat.EndConnect Method (Excel)

Attaches the end of the specified connector to a specified shape. If there?s already a connection between the end of the connector and another shape, that connection is broken. If the end of the connector isn?t already positioned at the specified connecting site, this method moves the end of the connector to the connecting site and adjusts the size and position of the connector. Use the  **[BeginConnect](connectorformat-beginconnect-method-excel.md)** method to attach the beginning of the connector to a shape.


## Syntax

 _expression_ . **EndConnect**( **_ConnectedShape_** , **_ConnectionSite_** )

 _expression_ A variable that represents a **ConnectorFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ConnectedShape_|Required| **[Shape](shape-object-excel.md)**|The shape to attach the end of the connector to. The specified  **Shape** object must be in the same **[Shapes](shapes-object-excel.md)** collection as the connector.|
| _ConnectionSite_|Required| **Long**|Must be an integer between 1 and the integer returned by the  **ConnectionSiteCount** property of the specified shape. If you want the connector to automatically find the shortest path between the two shapes it connects, specify any valid integer for this argument and then use the **[RerouteConnections](shape-rerouteconnections-method-excel.md)** method after the connector is attached to shapes at both ends.|

## Remarks

When you attach a connector to an object, the size and position of the connector are automatically adjusted, if necessary.


## Example

This example adds two rectangles to  `myDocument` and connects them with a curved connector. Notice that the **RerouteConnections** method makes it irrelevant what values you supply for the _ConnectionSite_ arguments used with the **BeginConnect** and **EndConnect** methods.


```vb
Set myDocument = Worksheets(1) 
Set s = myDocument.Shapes 
Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) 
Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) 
Set c = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100) 
With c.ConnectorFormat 
 .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1 
 .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1 
 c.RerouteConnections 
End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-excel.md)

