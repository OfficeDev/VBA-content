---
title: ConnectorFormat.EndConnect Method (PowerPoint)
keywords: vbapp10.chm555004
f1_keywords:
- vbapp10.chm555004
ms.prod: powerpoint
api_name:
- PowerPoint.ConnectorFormat.EndConnect
ms.assetid: b1a864e3-c2c2-ceeb-ac7c-5a26e7248dbe
ms.date: 06/08/2017
---


# ConnectorFormat.EndConnect Method (PowerPoint)

Attaches the end of the specified connector to a specified shape. 


## Syntax

 _expression_. **EndConnect**( **_ConnectedShape_**, **_ConnectionSite_** )

 _expression_ A variable that represents a **ConnectorFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ConnectedShape_|Required|**[Shape](shape-object-powerpoint.md)**|The shape to attach the end of the connector to. The specified  **Shape** object must be in the same **[Shapes](shapes-object-powerpoint.md)** collection as the connector.|
| _ConnectionSite_|Required|**Long**|A connection site on the shape specified by ConnectedShape. Must be an integer between 1 and the integer returned by the  **ConnectionSiteCount** property of the specified shape. If you want the connector to automatically find the shortest path between the two shapes it connects, specify any valid integer for this argument and then use the **RerouteConnections** method after the connector is attached to shapes at both ends.|

## Remarks

If there's already a connection between the end of the connector and another shape, that connection is broken. If the end of the connector isn't already positioned at the specified connecting site, this method moves the end of the connector to the connecting site and adjusts the size and position of the connector. Use the  **BeginConnect** method to attach the beginning of the connector to a shape.

When you attach a connector to an object, the size and position of the connector are automatically adjusted, if necessary.


## Example

This example adds two rectangles to the first slide in the active presentation and connects them with a curved connector. Notice that the  **RerouteConnections** method makes it irrelevant what values you supply for the ConnectionSite arguments used with the **BeginConnect** and **EndConnect** methods.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set s = myDocument.Shapes
Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)

With s.AddConnector(msoConnectorCurve, 0, 0, 100, 100) _
        .ConnectorFormat
    .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1
    .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1
    .Parent.RerouteConnections
End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-powerpoint.md)

