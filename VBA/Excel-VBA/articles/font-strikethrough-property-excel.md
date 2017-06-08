---
title: Font.Strikethrough Property (Excel)
keywords: vbaxl10.chm559083
f1_keywords:
- vbaxl10.chm559083
ms.prod: excel
api_name:
- Excel.Font.Strikethrough
ms.assetid: fc505f12-66ae-a941-c6cf-90f81bc44dea
ms.date: 06/08/2017
---


# Font.Strikethrough Property (Excel)

 **True** if the font is struck through with a horizontal line. Read/write **Boolean** .


## Syntax

 _expression_ . **Strikethrough**

 _expression_ A variable that represents a **Font** object.


## Example

This example sets the font in the active cell on Sheet1 to strikethrough.


```vb
Worksheets("Sheet1").Activate 
ActiveCell.Font.Strikethrough = True
```


## See also


#### Concepts


[Font Object](font-object-excel.md)

