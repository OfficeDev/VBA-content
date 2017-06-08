---
title: Border Object (Excel)
keywords: vbaxl10.chm546072
f1_keywords:
- vbaxl10.chm546072
ms.prod: excel
api_name:
- Excel.Border
ms.assetid: bca516bf-7c0f-f9df-078d-dfb522f256f3
ms.date: 06/08/2017
---


# Border Object (Excel)

Represents the border of an object.


## Remarks

Most bordered objects (all except for the  **[Range](range-object-excel.md)** and **[Style](style-object-excel.md)** objects) have a border that's treated as a single entity, regardless of how many sides it has. The entire border must be returned as a unit. Use the **[Border](trendline-border-property-excel.md)** property, such as from a **[TrendLine](trendline-object-excel.md)** object, to return the **Border** object for this kind of object.


## Example

 The following example changes the type and line style of a trend line on the active chart.


```
With ActiveChart.SeriesCollection(1).Trendlines(1) 
 .Type = xlLinear 
 .Border.LineStyle = xlDash 
End With
```

 **Range** and **Style** objects have four discrete borders — left, right, top, and bottom — which can be returned individually or as a group. Use the **Borders** property to return the **Borders** collection, which contains all four borders and treats the borders as a unit. The following example adds a double border to cell A1 on worksheet one.




```
Worksheets(1).Range("A1").Borders.LineStyle = xlDouble
```

Use  **Borders** ( _index_ ), where _index_ identifies the border, to return a single **Border** object. The following example sets the color of the bottom border of cells A1:G1.




```
Worksheets("Sheet1").Range("A1:G1"). _ 
 Borders(xlEdgeBottom).Color = RGB(255, 0, 0)
```

 _Index_ can be one of the following **[XlBordersIndex](xlbordersindex-enumeration-excel.md)** constants: **xlDiagonalDown**, **xlDiagonalUp**, **xlEdgeBottom**, **xlEdgeLeft**, **xlEdgeRight**, **xlEdgeTop**, **xlInsideHorizontal**, or **xlInsideVertical**.


## Properties



|**Name**|
|:-----|
|[Application](border-application-property-excel.md)|
|[Color](border-color-property-excel.md)|
|[ColorIndex](border-colorindex-property-excel.md)|
|[Creator](border-creator-property-excel.md)|
|[LineStyle](border-linestyle-property-excel.md)|
|[Parent](border-parent-property-excel.md)|
|[ThemeColor](border-themecolor-property-excel.md)|
|[TintAndShade](border-tintandshade-property-excel.md)|
|[Weight](border-weight-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
