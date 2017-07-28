---
title: Font.FontStyle Property (Excel)
keywords: vbaxl10.chm559077
f1_keywords:
- vbaxl10.chm559077
ms.prod: excel
api_name:
- Excel.Font.FontStyle
ms.assetid: 17e5989e-09a5-dabb-4989-82daf3aa0295
ms.date: 06/08/2017
---


# Font.FontStyle Property (Excel)

Returns or sets the font style. Read/write  **String** .


## Syntax

 _expression_ . **FontStyle**

 _expression_ A variable that represents a **Font** object.


## Remarks

Changing this property may affect other  **[Font](font-object-excel.md)** properties (such as **[Bold](texteffectformat-fontbold-property-excel.md)** and **[Italic](texteffectformat-fontitalic-property-excel.md)** ). Acceptable values are Regular, Italic, Bold, and Bold Italic.


## Example

This example sets the font style for cell A1 on Sheet1 to bold and italic.


```vb
Worksheets("Sheet1").Range("A1").Font.FontStyle = "Bold Italic"
```


## See also


#### Concepts


[Font Object](font-object-excel.md)

