---
title: ConnectorFormat.BeginConnectedShape Property (PowerPoint)
keywords: vbapp10.chm555007
f1_keywords:
- vbapp10.chm555007
ms.prod: powerpoint
api_name:
- PowerPoint.ConnectorFormat.BeginConnectedShape
ms.assetid: 7456899e-3f1c-3af8-e942-a6de1abeeca3
ms.date: 06/08/2017
---


# ConnectorFormat.BeginConnectedShape Property (PowerPoint)

Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the shape that the beginning of the specified connector is attached to. Read-only.


## Syntax

 _expression_. **BeginConnectedShape**

 _expression_ A variable that represents a **ConnectorFormat** object.


### Return Value

Shape


## Remarks

If the beginning of the specified connector isn't attached to a shape, this property generates an error.


## Example

This example assumes that the first slide in the active presentation already contains two shapes attached by a connector named "Conn1To2." The code adds a rectangle and a connector to the first slide. The beginning of the new connector will be attached to the same connection site as the beginning of the connector named "Conn1To2," and the end of the new connector will be attached to connection site one on the new rectangle.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes
    Set r3 = .AddShape(msoShapeRectangle, 450, 190, 200, 100)

    .AddConnector(msoConnectorCurve, 0, 0, 10, 10) _
        .Name = "Conn1To3"

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


[ConnectorFormat Object](connectorformat-object-powerpoint.md)

