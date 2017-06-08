---
title: HPageBreaks.Add Method (Excel)
keywords: vbaxl10.chm165076
f1_keywords:
- vbaxl10.chm165076
ms.prod: excel
api_name:
- Excel.HPageBreaks.Add
ms.assetid: 58aabcbf-7a9f-96a5-c91e-7311e397cffe
ms.date: 06/08/2017
---


# HPageBreaks.Add Method (Excel)

Adds a horizontal page break.


## Syntax

 _expression_ . **Add**( **_Before_** )

 _expression_ A variable that represents a **HPageBreaks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Before_|Required| **Object**|A  **[Range](range-object-excel.md)** object. The range above which the new page break will be added.|

### Return Value

A  **[HPageBreak](hpagebreak-object-excel.md)** object that represents the new horizontal page break.


## Example

This example adds a horizontal page break above cell F25 and adds a vertical page break to the left of this cell.


```vb
With Worksheets(1) 
 .HPageBreaks.Add .Range("F25") 
 .VPageBreaks.Add .Range("F25") 
End With
```


## See also


#### Concepts


[HPageBreaks Object](hpagebreaks-object-excel.md)

