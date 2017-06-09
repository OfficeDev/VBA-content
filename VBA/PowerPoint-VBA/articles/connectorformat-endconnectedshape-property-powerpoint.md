---
title: ConnectorFormat.EndConnectedShape Property (PowerPoint)
keywords: vbapp10.chm555010
f1_keywords:
- vbapp10.chm555010
ms.prod: powerpoint
api_name:
- PowerPoint.ConnectorFormat.EndConnectedShape
ms.assetid: 1d18fd9a-fc43-b50e-5c1a-dc6b5332b37e
ms.date: 06/08/2017
---


# ConnectorFormat.EndConnectedShape Property (PowerPoint)

Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the shape that the end of the specified connector is attached to. Read-only.


## Syntax

 _expression_. **EndConnectedShape**

 _expression_ A variable that represents an **ConnectorFormat** object.


### Return Value

Shape


## Remarks

If the end of the specified connector isn't attached to a shape, this property generates an error.


## Example

This example assumes that the fist slide in the active presentation already contains two shapes attached by a connector named "Conn1To2." The code adds a rectangle and a connector to the first slide. The end of the new connector will be attached to the same connection site as the end of the connector named "Conn1To2," and the beginning of the new connector will be attached to connection site one on the new rectangle.


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

