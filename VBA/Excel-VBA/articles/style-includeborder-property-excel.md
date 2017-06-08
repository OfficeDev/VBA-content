---
title: Style.IncludeBorder Property (Excel)
keywords: vbaxl10.chm177081
f1_keywords:
- vbaxl10.chm177081
ms.prod: excel
api_name:
- Excel.Style.IncludeBorder
ms.assetid: 81b44216-e8fa-88fe-e82c-7fd8844d33ea
ms.date: 06/08/2017
---


# Style.IncludeBorder Property (Excel)

 **True** if the style includes the **Color** , **ColorIndex** , **LineStyle** , and **Weight** border properties. Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeBorder**

 _expression_ A variable that represents a **Style** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include border format.


```vb
Worksheets("Sheet1").Range("A1").Style.IncludeBorder = True
```


## See also


#### Concepts


[Style Object](style-object-excel.md)

