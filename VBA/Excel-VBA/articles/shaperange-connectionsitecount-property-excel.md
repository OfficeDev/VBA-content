---
title: ShapeRange.ConnectionSiteCount Property (Excel)
keywords: vbaxl10.chm640100
f1_keywords:
- vbaxl10.chm640100
ms.prod: excel
api_name:
- Excel.ShapeRange.ConnectionSiteCount
ms.assetid: ce638d98-1db8-3f76-3f83-a38c62a04a1e
ms.date: 06/08/2017
---


# ShapeRange.ConnectionSiteCount Property (Excel)

Returns the number of connection sites on the specified shape. Read-only  **Long** .


## Syntax

 _expression_ . **ConnectionSiteCount**

 _expression_ An expression that returns a **ShapeRange** object.


## Example

This example adds two rectangles to  `myDocument` and joins them with two connectors. The beginnings of both connectors attach to connection site one on the first rectangle; the ends of the connectors attach to the first and last connection sites of the second rectangle.


```vb
Set myDocument = Worksheets(1) 
Set s = myDocument.Shapes 
Set firstRect = s.AddShape(msoShapeRectangle, _ 
 100, 50, 200, 100) 
Set secondRect = s.AddShape(msoShapeRectangle, _ 
 300, 300, 200, 100) 
lastsite = secondRect.ConnectionSiteCount 
With s.AddConnector(msoConnectorCurve, _ 
 0, 0, 100, 100).ConnectorFormat 
 .BeginConnect ConnectedShape:=firstRect, _ 
 ConnectionSite:=1 
 .EndConnect ConnectedShape:=secondRect, _ 
 ConnectionSite:=1 
End With 
With s.AddConnector(msoConnectorCurve, _ 
 0, 0, 100, 100).ConnectorFormat 
 .BeginConnect ConnectedShape:=firstRect, _ 
 ConnectionSite:=1 
 .EndConnect ConnectedShape:=secondRect, _ 
 ConnectionSite:=lastsite 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

