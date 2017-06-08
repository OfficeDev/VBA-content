---
title: ChartObjects.PrintObject Property (Excel)
keywords: vbaxl10.chm497087
f1_keywords:
- vbaxl10.chm497087
ms.prod: excel
api_name:
- Excel.ChartObjects.PrintObject
ms.assetid: 310a4571-e5e4-14c8-56a0-6d70a59f4588
ms.date: 06/08/2017
---


# ChartObjects.PrintObject Property (Excel)

 **True** if the objects will be printed when the document is printed. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintObject**

 _expression_ A variable that represents a **ChartObjects** object.


## Example

This example sets the objects on Sheet1 to be printed with the worksheet.


```vb
Worksheets("Sheet1").ChartObjects.PrintObject = True
```


