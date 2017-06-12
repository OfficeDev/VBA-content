---
title: ThreeDFormat.ExtrusionColorType Property (Excel)
keywords: vbaxl10.chm119007
f1_keywords:
- vbaxl10.chm119007
ms.prod: excel
api_name:
- Excel.ThreeDFormat.ExtrusionColorType
ms.assetid: cb463711-c8a3-5ac4-c81c-83d7b2d6b824
ms.date: 06/08/2017
---


# ThreeDFormat.ExtrusionColorType Property (Excel)

Returns or sets a value that indicates whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Read/write  **[MsoExtrusionColorType](http://msdn.microsoft.com/library/6acf7f2b-3d7b-15e3-f468-7dcb20865dc1%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **ExtrusionColorType**

 _expression_ A variable that represents a **ThreeDFormat** object.


## Remarks





| **MsoExtrusionColorType** can be one of these **MsoExtrusionColorType** constants.|
| **msoExtrusionColorAutomatic** . Extrusion color based on shape fill.|
| **msoExtrusionColorTypeMixed**|
| **msoExtrusionColorCustom** . Extrusion color independent of shape fill.|

## Example

If shape one on  `myDocument` has an automatic extrusion color, this example gives the extrusion a custom yellow color.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
    If .ExtrusionColorType = msoExtrusionColorAutomatic Then 
        .ExtrusionColor.RGB = RGB(240, 235, 16) 
    End If 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)

