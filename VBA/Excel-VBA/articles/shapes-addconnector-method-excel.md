---
title: Shapes.AddConnector Method (Excel)
keywords: vbaxl10.chm638078
f1_keywords:
- vbaxl10.chm638078
ms.prod: excel
api_name:
- Excel.Shapes.AddConnector
ms.assetid: 7ea648eb-ac6b-981d-652b-40cea1b3a8da
ms.date: 06/08/2017
---


# Shapes.AddConnector Method (Excel)

Creates a connector. Returns a  **[Shape](shape-object-excel.md)** object that represents the new connector. When a connector is added, it's not connected to anything. Use the **[BeginConnect](connectorformat-beginconnect-method-excel.md)** and **[EndConnect](connectorformat-endconnect-method-excel.md)** methods to attach the beginning and end of a connector to other shapes in the document.


## Syntax

 _expression_ . **AddConnector**( **_Type_** , **_BeginX_** , **_BeginY_** , **_EndX_** , **_EndY_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[MsoConnectorType](http://msdn.microsoft.com/library/2c67963f-5cb3-295d-fdf4-df33a283f1af%28Office.15%29.aspx)**|The connector type to add.|
| _BeginX_|Required| **Single**|The horizontal position (in points) of the connector's starting point relative to the upper-left corner of the document.|
| _BeginY_|Required| **Single**|The vertical position (in points) of the connector's starting point relative to the upper-left corner of the document.|
| _EndX_|Required| **Single**|The horizontal position (in points) of the connector's end point relative to the upper-left corner of the document.|
| _EndY_|Required| **Single**|The veritcal position (in points) of the connector's end point relative to the upper-left corner of the document.|

### Return Value

Shape


## Remarks



| **MsoConnectorType** can be one of these **MsoConnectorType** constants.|
| **msoConnectorElbow**|
| **msoConnectorTypeMixed**|
| **msoConnectorCurve**|
| **msoConnectorStraight**|
When you attach a connector to a shape, the size and position of the connector are automatically adjusted, if necessary. Therefore, if you?re going to attach a connector to other shapes, the position and dimensions you specify when adding the connector are irrelevant.


## Example

The following example adds a curved connector to a new canvas in a new worksheet.


```vb
Sub AddCanvasConnector() 
 
    Dim wksNew As Worksheet 
    Dim shpCanvas As Shape 
 
    Set wksNew = Worksheets.Add 
 
    'Add drawing canvas to new worksheet 
    Set shpCanvas = wksNew.Shapes.AddCanvas( _ 
        Left:=150, Top:=150, Width:=200, Height:=300) 
 
    'Add connector to the drawing canvas 
    shpCanvas.CanvasItems.AddConnector _ 
        Type:=msoConnectorStraight, BeginX:=150, _ 
        BeginY:=150, EndX:=200, EndY:=200 
 
End Sub
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

