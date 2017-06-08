---
title: Chart.PageSetup Property (Excel)
keywords: vbaxl10.chm148085
f1_keywords:
- vbaxl10.chm148085
ms.prod: excel
api_name:
- Excel.Chart.PageSetup
ms.assetid: 9a47bfd6-10b5-5f8e-86c2-e56c468de9d8
ms.date: 06/08/2017
---


# Chart.PageSetup Property (Excel)

Returns a  **[PageSetup](pagesetup-object-excel.md)** object that contains all the page setup settings for the specified object. Read-only.


## Syntax

 _expression_ . **PageSetup**

 _expression_ A variable that represents a **Chart** object.


## Example

This example sets the center header text for Chart1.


```vb
Charts("Chart1").PageSetup.CenterHeader = "December Sales"
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

