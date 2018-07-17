---
title: Application.MaxChange Property (Excel)
keywords: vbaxl10.chm133162
f1_keywords:
- vbaxl10.chm133162
ms.prod: excel
api_name:
- Excel.Application.MaxChange
ms.assetid: 5620bdff-d006-8c85-a1b8-1e3b31f21092
ms.date: 06/08/2017
---


# Application.MaxChange Property (Excel)

Returns or sets the maximum amount of change between each iteration as Microsoft Excel resolves circular references. Read/write  **Double** .


## Syntax

 _expression_ . **MaxChange**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **[MaxIterations](application-maxiterations-property-excel.md)** property sets the maximum number of iterations that Microsoft Excel can use when resolving circular references.


## Example

This example sets the maximum amount of change for each iteration to 0.1.


```vb
Application.MaxChange = 0.1
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

