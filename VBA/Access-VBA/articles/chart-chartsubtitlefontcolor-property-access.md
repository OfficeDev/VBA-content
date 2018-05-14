---
title: Chart.ChartSubtitleFontColor Property (Access)
keywords: vbaac10.chm6159
f1_keywords:
- vbaac10.chm6159
ms.prod: access
api_name:
- Access.Chart.ChartSubtitleFontColor
ms.date: 05/02/2018
---


# Chart.ChartSubtitleFontColor Property (Access)

Returns or sets the font color used by the chart subtitle. Read/write **Long** .

You can use a **[system color constant](../../language-reference-vba/articles/system-color-constants.md)** or the RGB function to set a color programmatically as shown in the example below. You can also browse and select a color from the Design View palette.


## Syntax

 _expression_ . **ChartSubtitleFontColor**

 _expression_ A variable that represents a **Chart** object.


## Example

In this example the **ChartSubtitleFontColor** is initially set to a system color constant before it is changed to an RGB value.
```vb
With myChart
 MsgBox ("Applying a system color constant")
 .ChartSubtitleFontColor = vbHighlight
 MsgBox ("Applying an RGB value")
 .ChartSubtitleFontColor = RGB(255, 165, 0)
End With
```

## See also


#### Concepts


[System Color Constants](../../language-reference-vba/articles/system-color-constants.md)

[Chart Object](chart-object-access.md)