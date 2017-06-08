---
title: Shapes.AddLine Method (Excel)
keywords: vbaxl10.chm638081
f1_keywords:
- vbaxl10.chm638081
ms.prod: excel
api_name:
- Excel.Shapes.AddLine
ms.assetid: f2186af3-c28a-4196-a00e-00cc66d28f71
ms.date: 06/08/2017
---


# Shapes.AddLine Method (Excel)

As it applies to the  **Shapes** object, returns a **[Shape](shape-object-excel.md)** object that represents the new line in a worksheet.


## Syntax

 _expression_ . **AddLine**( **_BeginX_** , **_BeginY_** , **_EndX_** , **_EndY_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BeginX_|Required| **Single**|The position (in points) of the line's starting point relative to the upper-left corner of the document.|
| _BeginY_|Required| **Single**|The position (in points) of the line's starting point relative to the upper-left corner of the document.|
| _EndX_|Required| **Single**|The position (in points) of the line's end point relative to the upper-left corner of the document.|
| _EndY_|Required| **Single**|The position (in points) of the line's end point relative to the upper-left corner of the document.|

### Return Value

Shape


## Example

This example adds a blue dashed line to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddLine(10, 10, 250, 250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

