---
title: ColorFormat Object (Word)
keywords: vbawd10.chm2502
f1_keywords:
- vbawd10.chm2502
ms.prod: word
api_name:
- Word.ColorFormat
ms.assetid: 5f12793f-d847-ecf2-6cf6-39387f7f0b28
ms.date: 06/08/2017
---


# ColorFormat Object (Word)

Represents the color of a one-color object or the foreground or background color of an object with a gradient or patterned fill. You can set colors to an explicit red-green-blue value by using the  **[RGB](colorformat-rgb-property-word.md)** property.


## Remarks

Use one of the properties listed in the following table to return a  **ColorFormat** object.



| <strong>Use this property</strong>                                                                                                                                          | <strong>With this object</strong>                                                                                                     | <strong>To return a ColorFormat object that represents this</strong> |
|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------|:---------------------------------------------------------------------|
| <strong><a href="fillformat-backcolor-property-word.md" data-raw-source="[BackColor](fillformat-backcolor-property-word.md)">BackColor</a></strong>                         | <strong><a href="fillformat-object-word.md" data-raw-source="[FillFormat](fillformat-object-word.md)">FillFormat</a></strong>         | Background fill color (used in a shaded or patterned fill)           |
| <strong><a href="fillformat-forecolor-property-word.md" data-raw-source="[ForeColor](fillformat-forecolor-property-word.md)">ForeColor</a></strong>                         | <strong><a href="fillformat-object-word.md" data-raw-source="[FillFormat](fillformat-object-word.md)">FillFormat</a></strong>         | Foreground fill color (or the fill color for a solid fill)           |
| <strong><a href="lineformat-backcolor-property-word.md" data-raw-source="[BackColor](lineformat-backcolor-property-word.md)">BackColor</a></strong>                         | <strong><a href="lineformat-object-word.md" data-raw-source="[LineFormat](lineformat-object-word.md)">LineFormat</a></strong>         | Background line color (used in a patterned line)                     |
| <strong><a href="lineformat-forecolor-property-word.md" data-raw-source="[ForeColor](lineformat-forecolor-property-word.md)">ForeColor</a></strong>                         | <strong><a href="lineformat-object-word.md" data-raw-source="[LineFormat](lineformat-object-word.md)">LineFormat</a></strong>         | Foreground line color (or the line color for a solid line)           |
| <strong><a href="shadowformat-forecolor-property-word.md" data-raw-source="[ForeColor](shadowformat-forecolor-property-word.md)">ForeColor</a></strong>                     | <strong><a href="shadowformat-object-word.md" data-raw-source="[ShadowFormat](shadowformat-object-word.md)">ShadowFormat</a></strong> | Shadow color                                                         |
| <strong><a href="threedformat-extrusioncolor-property-word.md" data-raw-source="[ExtrusionColor](threedformat-extrusioncolor-property-word.md)">ExtrusionColor</a></strong> | <strong><a href="threedformat-object-word.md" data-raw-source="[ThreeDFormat](threedformat-object-word.md)">ThreeDFormat</a></strong> | Color of the sides of an extruded object                             |

Use the  **RGB** property to set a color to an explicit red-green-blue value. The following example adds a rectangle to the active document and then sets the foreground color, background color, and gradient for the rectangle's fill.




```
With ActiveDocument.Shapes _ 
 .AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


## Properties



|**Name**|
|:-----|
|[Application](colorformat-application-property-word.md)|
|[Brightness](colorformat-brightness-property-word.md)|
|[Creator](colorformat-creator-property-word.md)|
|[ObjectThemeColor](colorformat-objectthemecolor-property-word.md)|
|[Parent](colorformat-parent-property-word.md)|
|[RGB](colorformat-rgb-property-word.md)|
|[TintAndShade](colorformat-tintandshade-property-word.md)|
|[Type](colorformat-type-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
