---
title: Style.IncludeNumber Property (Excel)
keywords: vbaxl10.chm177083
f1_keywords:
- vbaxl10.chm177083
ms.prod: excel
api_name:
- Excel.Style.IncludeNumber
ms.assetid: bd46ac34-67bb-cb78-1ad6-321fc4210f84
ms.date: 06/08/2017
---


# Style.IncludeNumber Property (Excel)

 **True** if the style includes the **NumberFormat** property. Read/write **Boolean**


## Syntax

 _expression_ . **IncludeNumber**

 _expression_ A variable that represents a **Style** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include number format.


```vb
Worksheets("Sheet1").Range("A1").Style.IncludeNumber = True
```


## See also


#### Concepts


[Style Object](style-object-excel.md)

