---
title: Chart.PrimaryValuesAxisFontColor Property (Access)
keywords: vbaac10.chm6132
f1_keywords:
- vbaac10.chm6132
ms.prod: access
api_name:
- Access.Chart.PrimaryValuesAxisFontColor
ms.date: 05/02/2018
---


# Chart.PrimaryValuesAxisFontColor Property (Access)

Returns or sets the font color used by the primary values axis. Read/write **Long** .

You can use a **[system color constant](../../language-reference-vba/articles/system-color-constants.md)** or the RGB function to set a color programmatically as shown in the example below. You can also browse and select a color from the Design View palette.


## Syntax

 _expression_ . **PrimaryValuesAxisFontColor**

 _expression_ A variable that represents a **Chart** object.


## Example

In this example the **PrimaryValuesAxisFontColor** is initially set to a system color constant before it is changed to an RGB value.
```vb
With myChart
 MsgBox ("Applying a system color constant")
 .PrimaryValuesAxisFontColor = vbHighlight
 MsgBox ("Applying an RGB value")
 .PrimaryValuesAxisFontColor = RGB(255, 165, 0)
End With
```

## See also


#### Concepts


[System Color Constants](../../language-reference-vba/articles/system-color-constants.md)

[Chart Object](chart-object-access.md)