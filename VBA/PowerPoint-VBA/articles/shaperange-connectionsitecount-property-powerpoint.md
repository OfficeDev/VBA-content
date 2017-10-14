---
title: ShapeRange.ConnectionSiteCount Property (PowerPoint)
keywords: vbapp10.chm548019
f1_keywords:
- vbapp10.chm548019
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.ConnectionSiteCount
ms.assetid: 352f9c7c-6290-f974-5924-01e108fb4919
ms.date: 06/08/2017
---


# ShapeRange.ConnectionSiteCount Property (PowerPoint)

Returns the number of connection sites on the specified shape. Read-only.


## Syntax

 _expression_. **ConnectionSiteCount**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

Long


## Example

This example adds two rectangles to  `myDocument` and joins them with two connectors. The beginnings of both connectors attach to connection site one on the first rectangle; the ends of the connectors attach to the first and last connection sites of the second rectangle.


```vb
Set myDocument = ActivePresentation.Slides(1)
Set s = myDocument.Shapes
Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)

lastsite = secondRect.ConnectionSiteCount

With s.AddConnector(msoConnectorCurve, 0, 0, 100, 100) _
        .ConnectorFormat

    .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1
    .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1

End With

With s.AddConnector(msoConnectorCurve, 0, 0, 100, 100) _
        .ConnectorFormat

    .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1
    .EndConnect ConnectedShape:=secondRect, _
        ConnectionSite:=lastsite

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

