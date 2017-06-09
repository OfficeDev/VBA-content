---
title: Application.MaxIterations Property (Excel)
keywords: vbaxl10.chm133163
f1_keywords:
- vbaxl10.chm133163
ms.prod: excel
api_name:
- Excel.Application.MaxIterations
ms.assetid: 83f12597-9186-e415-a22b-9e028bd95169
ms.date: 06/08/2017
---


# Application.MaxIterations Property (Excel)

Returns or sets the maximum number of iterations that Microsoft Excel can use to resolve a circular reference. Read/write  **Long** .


## Syntax

 _expression_ . **MaxIterations**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **[MaxChange](application-maxchange-property-excel.md)** property sets the maximum amount of change between each iteration when Microsoft Excel is resolving circular references.


## Example

This example sets the maximum number of iterations at 1000.


```vb
Application.MaxIterations = 1000
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

