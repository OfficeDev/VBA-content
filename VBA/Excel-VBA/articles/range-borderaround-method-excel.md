---
title: Range.BorderAround Method (Excel)
keywords: vbaxl10.chm144252
f1_keywords:
- vbaxl10.chm144252
ms.prod: excel
api_name:
- Excel.Range.BorderAround
ms.assetid: 3ffeb131-45f7-7799-e04a-11577fedaa16
ms.date: 06/08/2017
---


# Range.BorderAround Method (Excel)

Adds a border to a range and sets the  **[Color](border-color-property-excel.md)** , **[LineStyle](border-linestyle-property-excel.md)** , and **[Weight](border-weight-property-excel.md)** properties for the new border. **Variant** .


## Syntax

 _expression_ . **BorderAround**( **_LineStyle_** , **_Weight_** , **_ColorIndex_** , **_Color_** , **_ThemeColor_** )

 _expression_ A variable that represents a **[Range](range-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LineStyle_|Optional| **Variant**|One of the constants of  **[XlLineStyle](xllinestyle-enumeration-excel.md)** specifying the line style for the border.|
| _Weight_|Optional| **[XlBorderWeight](xlborderweight-enumeration-excel.md)**|The border weight.|
| _ColorIndex_|Optional| **[XlColorIndex](xlcolorindex-enumeration-excel.md)**|The border color, as an index into the current color palette or as an  **XlColorIndex** constant.|
| _Color_|Optional| **Variant**|The border color, as an RGB value.|
| _ThemeColor_|Optional| **Variant**|The theme color, as an index into the current color theme or as an  **[XlThemeColor](xlthemecolor-enumeration-excel.md)** value.|

### Return Value

Variant


## Remarks

You must specify only one of the following:  _ColorIndex_,  _Color_, or  _ThemeColor_.

You can specify either  _LineStyle_ or _Weight_, but not both. If you don't specify either argument, Microsoft Excel uses the default line style and weight.

This method outlines the entire range without filling it in. To set the borders of all the cells, you must set the  **Color** , **LineStyle** , and **Weight** properties for the **[Borders](borders-object-excel.md)** collection. To clear the border, you must set the **LineStyle** property to **xlLineStyleNone** for all the cells in the range.


## Example

This example adds a thick red border around the range A1:D4 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1:D4").BorderAround _ 
 ColorIndex:=3, Weight:=xlThick
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

