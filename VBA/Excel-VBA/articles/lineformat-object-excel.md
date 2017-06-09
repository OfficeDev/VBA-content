---
title: LineFormat Object (Excel)
keywords: vbaxl10.chm110000
f1_keywords:
- vbaxl10.chm110000
ms.prod: excel
api_name:
- Excel.LineFormat
ms.assetid: 13eca34b-adf7-ddd3-8c73-cc8b508c624a
ms.date: 06/08/2017
---


# LineFormat Object (Excel)

Represents line and arrowhead formatting.


## Remarks

 For a line, the **LineFormat** object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.


## Example

Use the  **[Line](shape-line-property-excel.md)** property to return a **LineFormat** object. The following example adds a blue, dashed line to _myDocument_. There's a short, narrow oval at the line's starting point and a long, wide triangle at its end point.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddLine(100, 100, 200, 300).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
End With
```


## Properties



|**Name**|
|:-----|
|[Application](lineformat-application-property-excel.md)|
|[BackColor](lineformat-backcolor-property-excel.md)|
|[BeginArrowheadLength](lineformat-beginarrowheadlength-property-excel.md)|
|[BeginArrowheadStyle](lineformat-beginarrowheadstyle-property-excel.md)|
|[BeginArrowheadWidth](lineformat-beginarrowheadwidth-property-excel.md)|
|[Creator](lineformat-creator-property-excel.md)|
|[DashStyle](lineformat-dashstyle-property-excel.md)|
|[EndArrowheadLength](lineformat-endarrowheadlength-property-excel.md)|
|[EndArrowheadStyle](lineformat-endarrowheadstyle-property-excel.md)|
|[EndArrowheadWidth](lineformat-endarrowheadwidth-property-excel.md)|
|[ForeColor](lineformat-forecolor-property-excel.md)|
|[InsetPen](lineformat-insetpen-property-excel.md)|
|[Parent](lineformat-parent-property-excel.md)|
|[Pattern](lineformat-pattern-property-excel.md)|
|[Style](lineformat-style-property-excel.md)|
|[Transparency](lineformat-transparency-property-excel.md)|
|[Visible](lineformat-visible-property-excel.md)|
|[Weight](lineformat-weight-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
