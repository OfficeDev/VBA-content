---
title: Font Object (Excel)
keywords: vbaxl10.chm558072
f1_keywords:
- vbaxl10.chm558072
ms.prod: excel
api_name:
- Excel.Font
ms.assetid: f4788ba4-1c4c-2f03-4d73-194bc9316825
ms.date: 06/08/2017
---


# Font Object (Excel)

Contains the font attributes (font name, font size, color, and so on) for an object.


## Remarks

If you don't want to format all the text in a cell or graphic the same way, use the  **[Characters](range-characters-property-excel.md)** property to return a subset of the text.


## Example

Use the  **Font** property to return the **Font** object. The following example formats cells A1:C5 as bold.


```
Worksheets("Sheet1").Range("A1:C5").Font.Bold = True
```


## Properties



|**Name**|
|:-----|
|[Application](font-application-property-excel.md)|
|[Background](font-background-property-excel.md)|
|[Bold](font-bold-property-excel.md)|
|[Color](font-color-property-excel.md)|
|[ColorIndex](font-colorindex-property-excel.md)|
|[Creator](font-creator-property-excel.md)|
|[FontStyle](font-fontstyle-property-excel.md)|
|[Italic](font-italic-property-excel.md)|
|[Name](font-name-property-excel.md)|
|[Parent](font-parent-property-excel.md)|
|[Size](font-size-property-excel.md)|
|[Strikethrough](font-strikethrough-property-excel.md)|
|[Subscript](font-subscript-property-excel.md)|
|[Superscript](font-superscript-property-excel.md)|
|[ThemeColor](font-themecolor-property-excel.md)|
|[ThemeFont](font-themefont-property-excel.md)|
|[TintAndShade](font-tintandshade-property-excel.md)|
|[Underline](font-underline-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
