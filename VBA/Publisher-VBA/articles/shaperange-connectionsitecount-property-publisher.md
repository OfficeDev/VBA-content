---
title: ShapeRange.ConnectionSiteCount Property (Publisher)
keywords: vbapb10.chm2293812
f1_keywords:
- vbapb10.chm2293812
ms.prod: publisher
api_name:
- Publisher.ShapeRange.ConnectionSiteCount
ms.assetid: f830739d-08be-562c-83fc-7f7a6f8e047c
ms.date: 06/08/2017
---


# ShapeRange.ConnectionSiteCount Property (Publisher)

Returns a  **Long** indicating the count of connection sites on the current **Shape** object. Read-only.


## Syntax

 _expression_. **ConnectionSiteCount**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks

The number of connection sites varies depending on the shape geometry. Rectangular objects including tables and Web controls will most likely have four connection sites, one centered on each edge of the shape.


## Example

This example adds two rectangles to the active publication and joins them with two connectors. The beginnings of both connectors attach to connection site one on the first rectangle; the ends of the connectors attach to the first and last connection sites of the second rectangle. Then it counts the number of connections on the first rectangle.


```vb
Sub Connections() 
 
 Dim shpNew As Shapes 
 Dim shpFirstRect As Shape 
 Dim shpSecondRect As Shape 
 Dim intLastSite As Integer 
 Dim intCount As Integer 
 
 Set shpNew = Application.ActiveDocument _ 
 .MasterPages(Item:=1).Shapes 
 Set shpFirstRect = shpNew.AddShape(Type:=msoShapeRectangle, _ 
 Left:=100, Top:=50, Width:=200, Height:=100) 
 Set shpSecondRect = shpNew.AddShape(msoShapeRectangle, _ 
 Left:=300, Top:=300, Width:=200, Height:=100) 
 varLastSite = shpSecondRect.ConnectionSiteCount 
 
 ' Add the first connector from rectangle 1, 
 ' site 1 to rectangle 2, site 1. 
 With shpNew.AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=100, EndY:=100) _ 
 .ConnectorFormat 
 .BeginConnect ConnectedShape:=shpFirstRect, ConnectionSite:=1 
 .EndConnect ConnectedShape:=shpSecondRect, ConnectionSite:=1 
 End With 
 
 ' Add the second connector from rectangle 1, 
 ' site 1 to rectangle 2, site 2. 
 With shpNew.AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=100, EndY:=100) _ 
 .ConnectorFormat 
 .BeginConnect ConnectedShape:=shpFirstRect, ConnectionSite:=1 
 .EndConnect ConnectedShape:=shpSecondRect, _ 
 ConnectionSite:=intLastSite 
 End With 
 
 intCount = shpFirstRect.ConnectionSiteCount 
 
End Sub
```


