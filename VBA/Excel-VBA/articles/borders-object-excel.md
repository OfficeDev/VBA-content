---
title: Borders Object (Excel)
keywords: vbaxl10.chm180072
f1_keywords:
- vbaxl10.chm180072
ms.prod: excel
api_name:
- Excel.Borders
ms.assetid: adb6efd6-73b6-e620-e9be-f4a42bc52ae8
ms.date: 06/08/2017
---


# Borders Object (Excel)

A collection of four  **[Border](border-object-excel.md)** objects that represent the four borders of a **[Range](range-object-excel.md)** or **[Style](style-object-excel.md)** object.


## Remarks

Use the  **Borders** property to return the **Borders** collection, which contains all four borders.

You can set border properties for an individual border only with  **Range** and **Style** objects. Other bordered objects, such as error bars and series lines, have a border that's treated as a single entity, regardless of how many sides it has. For these objects, you must return and set properties for the entire border as a unit. For more information, see the **Border** object.


## Example

The following example adds a double border to cell A1 on worksheet one.


```
Worksheets(1).Range("A1").Borders.LineStyle = xlDouble
```

Use  **Borders** ( _index_ ), where _index_ identifies the border, to return a single **Border** object. The following example sets the color of the bottom border of cells A1:G1 to red.




```
Worksheets("Sheet1").Range("A1:G1"). _ 
 Borders(xlEdgeBottom).Color = RGB(255, 0, 0)
```

 _Index_ can be one of the following **[xlBordersIndex](xlbordersindex-enumeration-excel.md)** constants: **xlDiagonalDown**, **xlDiagonalUp**, **xlEdgeBottom**, **xlEdgeLeft**, **xlEdgeRight**, or **xlEdgeTop**, **xlInsideHorizontal**, or **xlInsideVertical**.


## Properties



|**Name**|
|:-----|
|[Application](borders-application-property-excel.md)|
|[Color](borders-color-property-excel.md)|
|[ColorIndex](borders-colorindex-property-excel.md)|
|[Count](borders-count-property-excel.md)|
|[Creator](borders-creator-property-excel.md)|
|[Item](borders-item-property-excel.md)|
|[LineStyle](borders-linestyle-property-excel.md)|
|[Parent](borders-parent-property-excel.md)|
|[ThemeColor](borders-themecolor-property-excel.md)|
|[TintAndShade](borders-tintandshade-property-excel.md)|
|[Value](borders-value-property-excel.md)|
|[Weight](borders-weight-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
