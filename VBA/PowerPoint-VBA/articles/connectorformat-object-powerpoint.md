---
title: ConnectorFormat Object (PowerPoint)
keywords: vbapp10.chm555000
f1_keywords:
- vbapp10.chm555000
ms.prod: powerpoint
api_name:
- PowerPoint.ConnectorFormat
ms.assetid: 54504fab-8279-1012-db7f-3f19a4840637
ms.date: 06/08/2017
---


# ConnectorFormat Object (PowerPoint)

Contains properties and methods that apply to connectors. 


## Remarks

A connector is a line that attaches two other shapes at points called connection sites. If you rearrange shapes that are connected, the geometry of the connector will be automatically adjusted so that the shapes remain connected.

Use the  **ConnectorFormat** property to return a **ConnectorFormat** object. Use the **[BeginConnect](connectorformat-beginconnect-method-powerpoint.md)** and **[EndConnect](connectorformat-endconnect-method-powerpoint.md)** methods to attach the ends of the connector to other shapes in the document. Use the **[RerouteConnections](shaperange-rerouteconnections-method-powerpoint.md)** method to automatically find the shortest path between the two shapes connected by the connector. Use the **[Connector](shaperange-connector-property-powerpoint.md)** property to see whether a shape is a connector.

Note that you assign a size and a position when you add a connector to the  **Shapes** collection, but the size and position are automatically adjusted when you attach the beginning and end of the connector to other shapes in the collection. Therefore, if you intend to attach a connector to other shapes, the initial size and position you specify are irrelevant. Likewise, you specify which connection sites on a shape to attach the connector to when you attach the connector, but using the **RerouteConnections** method after the connector is attached may change which connection sites the connector attaches to, making your original choice of connection sites irrelevant.

Connection sites are generally numbered according to the rules presented in the following table.



|**Shape type**|**Connection site numbering scheme**|
|:-----|:-----|
|AutoShapes, WordArt, pictures, and OLE objects|The connection sites are numbered starting at the top and proceeding counterclockwise.|
|Freeforms|The connection sites are the vertices, and they correspond to the vertex numbers.|

## Example

The following example adds two rectangles to  `myDocument` and connects them with a curved connector.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set s = myDocument.Shapes

Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)

Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)

With s.AddConnector(msoConnectorCurve, 0, 0, 0, 0).ConnectorFormat

    .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1

    .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1

    .Parent.RerouteConnections

End With
```

To figure out which number corresponds to which connection site on a complex shape, you can experiment with the shape while the macro recorder is turned on and then examine the recorded code; or you can create a shape, select it, and then run the following example. This code will number each connection site and attach a connector to it.




```vb
Set mainshape = ActiveWindow.Selection.ShapeRange(1)

With mainshape

    bx = .Left + .Width + 50

    by = .Top + .Height + 50

End With

With ActiveWindow.View.Slide

    For j = 1 To mainshape.ConnectionSiteCount

        With .Shapes.AddConnector(msoConnectorStraight, bx, by, bx + 50, by + 50)

            .ConnectorFormat.EndConnect mainshape, j

            .ConnectorFormat.Type = msoConnectorElbow

            .Line.ForeColor.RGB = RGB(255, 0, 0)

            l = .Left

            t = .Top

        End With

        With .Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, 36, 14)

            .Fill.Visible = False

            .Line.Visible = False

            .TextFrame.TextRange.Text = j

        End With

    Next j

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

