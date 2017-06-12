---
title: Font.Italic Property (Excel)
keywords: vbaxl10.chm559078
f1_keywords:
- vbaxl10.chm559078
ms.prod: excel
api_name:
- Excel.Font.Italic
ms.assetid: 9d249157-9c8a-79ec-9b70-021c19ea1336
ms.date: 06/08/2017
---


# Font.Italic Property (Excel)

 **True** if the font style is italic. Read/write **Boolean** .


## Syntax

 _expression_ . **Italic**

 _expression_ A variable that represents a **Font** object.


## Example

This example sets the font style to italic for the range A1:A5 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1:A5").Font.Italic = True
```


## See also


#### Concepts


[Font Object](font-object-excel.md)

