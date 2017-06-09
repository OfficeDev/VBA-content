---
title: Borders.LineStyle Property (Excel)
keywords: vbaxl10.chm181077
f1_keywords:
- vbaxl10.chm181077
ms.prod: excel
api_name:
- Excel.Borders.LineStyle
ms.assetid: a057234d-0442-3fd7-5547-b19451774c0e
ms.date: 06/08/2017
---


# Borders.LineStyle Property (Excel)

Returns or sets the line style for the border. Read/write  **[XlLineStyle](xllinestyle-enumeration-excel.md)** , **xlGray25** , **xlGray50** , **xlGray75** , or **xlAutomatic** .


## Syntax

 _expression_ . **LineStyle**

 _expression_ A variable that represents a **Borders** object.


## Example

This example puts a border around the chart area and the plot area of Chart1.


```vb
With Charts("Chart1") 
 .ChartArea.Border.LineStyle = xlDashDot 
 With .PlotArea.Border 
 .LineStyle = xlDashDotDot 
 .Weight = xlThick 
 End With 
End With
```


## See also


#### Concepts


[Borders Collection](borders-object-excel.md)

