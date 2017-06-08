---
title: Range.SetPhonetic Method (Excel)
keywords: vbaxl10.chm144230
f1_keywords:
- vbaxl10.chm144230
ms.prod: excel
api_name:
- Excel.Range.SetPhonetic
ms.assetid: 69a1e491-5505-621a-5ea0-b0600796caa3
ms.date: 06/08/2017
---


# Range.SetPhonetic Method (Excel)

Creates  **[Phonetic](phonetic-object-excel.md)** objects for all the cells in the specified range.


## Syntax

 _expression_ . **SetPhonetic**

 _expression_ A variable that represents a **Range** object.


## Remarks

Any existing  **Phonetic** objects in the specified range are automatically overwritten (deleted) by the new objects you add using this method.


## Example

This example creates a  **Phonetic** object for each cell in the range A1:A10 on the active worksheet.


```vb
ActiveSheet.Range("A1:A10").SetPhonetic
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

