---
title: Style.IncludeFont Property (Excel)
keywords: vbaxl10.chm177082
f1_keywords:
- vbaxl10.chm177082
ms.prod: excel
api_name:
- Excel.Style.IncludeFont
ms.assetid: 280f866f-dcd8-dabd-0673-a26090e7f53a
ms.date: 06/08/2017
---


# Style.IncludeFont Property (Excel)

 **True** if the style includes the **[Background](font-background-property-excel.md)** , **[Bold](texteffectformat-fontbold-property-excel.md)** , **[Color](font-color-property-excel.md)** , **[ColorIndex](font-colorindex-property-excel.md)** , **[FontStyle](font-fontstyle-property-excel.md)** , **[Italic](texteffectformat-fontitalic-property-excel.md)** , **[Name](texteffectformat-fontname-property-excel.md)** , **[Size](texteffectformat-fontsize-property-excel.md)** , **[Strikethrough](font-strikethrough-property-excel.md)** , **[Subscript](font-subscript-property-excel.md)** , **[Superscript](font-superscript-property-excel.md)** , and **[Underline](font-underline-property-excel.md)** font properties. Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeFont**

 _expression_ A variable that represents a **Style** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include font format.


```vb
Worksheets("Sheet1").Range("A1").Style.IncludeFont = True
```


## See also


#### Concepts


[Style Object](style-object-excel.md)

