---
title: Border.LineStyle Property (Excel)
keywords: vbaxl10.chm547075
f1_keywords:
- vbaxl10.chm547075
ms.prod: excel
api_name:
- Excel.Border.LineStyle
ms.assetid: 7f2529b7-4782-8d8d-d529-6d8d19417db4
ms.date: 06/08/2017
---


# Border.LineStyle Property (Excel)

Returns or sets the line style for the border. Read/write  **[XlLineStyle](xllinestyle-enumeration-excel.md)** , **xlGray25** , **xlGray50** , **xlGray75** , or **xlAutomatic** .


## Syntax

 _expression_ . **LineStyle**

 _expression_ A variable that represents a **Border** object.


## Remarks

 **xlDouble** and **xlSlantDashDot** do not apply to charts.


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


[Border Object](border-object-excel.md)

