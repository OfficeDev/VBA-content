---
title: ColorFormat Object (Excel)
keywords: vbaxl10.chm105000
f1_keywords:
- vbaxl10.chm105000
ms.prod: excel
api_name:
- Excel.ColorFormat
ms.assetid: 9bb6bc1f-9886-d290-a336-068f84cad1a9
ms.date: 06/08/2017
---


# ColorFormat Object (Excel)

Represents the color of a one-color object, the foreground or background color of an object with a gradient or patterned fill, or the pointer color.


## Remarks

 You can set colors to an explicit red-green-blue value (by using the **[RGB](colorformat-rgb-property-excel.md)** property) or to a color in the color scheme (by using the **[SchemeColor](colorformat-schemecolor-property-excel.md)** property).

Use one of the properties listed in the following table to return a  **ColorFormat** object.



|**Use this property**|**With this object**|**To return a ColorFormat object that represents this**|
|:-----|:-----|:-----|
|**[BackColor](fillformat-backcolor-property-excel.md)**|**FillFormat**|The background fill color (used in a shaded or patterned fill)|
|**[ForeColor](fillformat-forecolor-property-excel.md)**|**FillFormat**|The foreground fill color (or simply the fill color for a solid fill)|
|**[BackColor](lineformat-backcolor-property-excel.md)**|**LineFormat**|The background line color (used in a patterned line)|
|**[ForeColor](lineformat-forecolor-property-excel.md)**|**LineFormat**|The foreground line color (or just the line color for a solid line)|
|**[ForeColor](shadowformat-forecolor-property-excel.md)**|**ShadowFormat**|The shadow color|
|**[ExtrusionColor](threedformat-extrusioncolor-property-excel.md)**|**ThreeDFormat**|The color of the sides of an extruded object|

## Example

Use the  **RGB** property to set a color to an explicit red-green-blue value. The following example adds a rectangle to _myDocument_ and then sets the foreground color, background color, and gradient for the rectangle's fill.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 90, 90, 90, 50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


## Properties



|**Name**|
|:-----|
|[Application](colorformat-application-property-excel.md)|
|[Brightness](colorformat-brightness-property-excel.md)|
|[Creator](colorformat-creator-property-excel.md)|
|[ObjectThemeColor](colorformat-objectthemecolor-property-excel.md)|
|[Parent](colorformat-parent-property-excel.md)|
|[RGB](colorformat-rgb-property-excel.md)|
|[SchemeColor](colorformat-schemecolor-property-excel.md)|
|[TintAndShade](colorformat-tintandshade-property-excel.md)|
|[Type](colorformat-type-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
