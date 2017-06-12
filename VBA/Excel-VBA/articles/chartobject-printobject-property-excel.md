---
title: ChartObject.PrintObject Property (Excel)
keywords: vbaxl10.chm494089
f1_keywords:
- vbaxl10.chm494089
ms.prod: excel
api_name:
- Excel.ChartObject.PrintObject
ms.assetid: 504f4a82-6129-cb38-ea2f-f9b29e14d036
ms.date: 06/08/2017
---


# ChartObject.PrintObject Property (Excel)

 **True** if the object will be printed when the document is printed. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintObject**

 _expression_ A variable that represents a **ChartObject** object.


## Example

This example sets embedded chart one on Sheet1 to be printed with the worksheet.


```vb
Worksheets("Sheet1").ChartObjects(1).PrintObject = True
```


## See also


#### Concepts


[ChartObject Object](chartobject-object-excel.md)

