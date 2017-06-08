---
title: Range.Speak Method (Excel)
keywords: vbaxl10.chm144237
f1_keywords:
- vbaxl10.chm144237
ms.prod: excel
api_name:
- Excel.Range.Speak
ms.assetid: 12928814-9534-c9f0-e351-7d26f77869e0
ms.date: 06/08/2017
---


# Range.Speak Method (Excel)

Causes the cells of the range to be spoken in row order or column order.


## Syntax

 _expression_ . **Speak**( **_SpeakDirection_** , **_SpeakFormulas_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SpeakDirection_|Optional| **Variant**|The speak direction, by rows or columns.|
| _SpeakFormulas_|Optional| **Variant**| **True** will cause formulas to be sent to the Text-To-Speech (TTS) engine for cells that have formulas. The value is sent if the cells do not have formulas. **False** (default) will cause values to always be sent to the TTS engine.|

## See also


#### Concepts


[Range Object](range-object-excel.md)

