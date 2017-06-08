---
title: Range.Validation Property (Excel)
keywords: vbaxl10.chm144215
f1_keywords:
- vbaxl10.chm144215
ms.prod: excel
api_name:
- Excel.Range.Validation
ms.assetid: d1cad7e6-bbfa-e280-33e7-048733efc0bc
ms.date: 06/08/2017
---


# Range.Validation Property (Excel)

Returns the  **[Validation](validation-object-excel.md)** object that represents data validation for the specified range. Read-only.


## Syntax

 _expression_ . **Validation**

 _expression_ A variable that represents a **Range** object.


## Example

This example causes data validation for cell E5 to allow blank values.


```vb
Range("e5").Validation.IgnoreBlank = True
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

