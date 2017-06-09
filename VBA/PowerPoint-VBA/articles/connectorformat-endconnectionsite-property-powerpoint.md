---
title: ConnectorFormat.EndConnectionSite Property (PowerPoint)
keywords: vbapp10.chm555011
f1_keywords:
- vbapp10.chm555011
ms.prod: powerpoint
api_name:
- PowerPoint.ConnectorFormat.EndConnectionSite
ms.assetid: fa65a404-573a-939b-6e2c-d54e4de5c1f0
ms.date: 06/08/2017
---


# ConnectorFormat.EndConnectionSite Property (PowerPoint)

Returns an integer that specifies the connection site that the end of a connector is connected to. Read-only. 


## Syntax

 _expression_. **EndConnectionSite**

 _expression_ A variable that represents an **ConnectorFormat** object.


### Return Value

Long


## Remarks

If the end of the specified connector isn't attached to a shape, this property generates an error.


## Example

This example assumes that the first slide in the active presentation already contains two shapes attached by a connector named "Conn1To2." The code adds a rectangle and a connector to the first slide. The end of the new connector will be attached to the same connection site as the end of the connector named "Conn1To2," and the beginning of the new connector will be attached to connection site one on the new rectangle.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes
    Set r3 = .AddShape(msoShapeRectangle, 100, 420, 200, 100)
    .AddConnector(msoConnectorCurve, 0, 0, 10, 10) _
        .Name = "Conn1To3"
    With .Item("Conn1To2").ConnectorFormat
        endConnSite1 = .EndConnectionSite
        Set endConnShape1 = .EndConnectedShape
    End With

    With .Item("Conn1To3").ConnectorFormat
        .BeginConnect r3, 1
        .EndConnect endConnShape1, endConnSite1
    End With
End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-powerpoint.md)

