---
title: PageSetup.CenterHorizontally Property (Excel)
keywords: vbaxl10.chm473077
f1_keywords:
- vbaxl10.chm473077
ms.prod: excel
api_name:
- Excel.PageSetup.CenterHorizontally
ms.assetid: 6b3e97fd-6b05-6863-c642-b085ea9ddd33
ms.date: 06/08/2017
---


# PageSetup.CenterHorizontally Property (Excel)

 **True** if the sheet is centered horizontally on the page when it's printed. Read/write **Boolean** .


## Syntax

 _expression_ . **CenterHorizontally**

 _expression_ A variable that represents a **PageSetup** object.


## Example

This example centers Sheet1 horizontally when it's printed.


```vb
Worksheets("Sheet1").PageSetup.CenterHorizontally = True
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

