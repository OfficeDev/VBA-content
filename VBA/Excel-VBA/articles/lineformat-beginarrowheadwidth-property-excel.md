---
title: LineFormat.BeginArrowheadWidth Property (Excel)
keywords: vbaxl10.chm110005
f1_keywords:
- vbaxl10.chm110005
ms.prod: excel
api_name:
- Excel.LineFormat.BeginArrowheadWidth
ms.assetid: 82d9b8fe-4aa5-3292-f792-c14332c2103d
ms.date: 06/08/2017
---


# LineFormat.BeginArrowheadWidth Property (Excel)

Returns or sets the width of the arrowhead at the beginning of the specified line. Read/write  **[MsoArrowheadWidth](http://msdn.microsoft.com/library/7183f2e0-7431-170b-f4e7-3f8737017ed8%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **BeginArrowheadWidth**

 _expression_ A variable that represents a **LineFormat** object.


## Example

This example adds a line to  `myDocument`. There's a short, narrow oval on the line's starting point and a long, wide triangle on its end point.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddLine(100, 100, 200, 300).Line 
    .BeginArrowheadLength = msoArrowheadShort 
    .BeginArrowheadStyle = msoArrowheadOval 
    .BeginArrowheadWidth = msoArrowheadNarrow 
    .EndArrowheadLength = msoArrowheadLong 
    .EndArrowheadStyle = msoArrowheadTriangle 
    .EndArrowheadWidth = msoArrowheadWide 
End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-excel.md)

