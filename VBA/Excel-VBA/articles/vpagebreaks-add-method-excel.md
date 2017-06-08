---
title: VPageBreaks.Add Method (Excel)
keywords: vbaxl10.chm168076
f1_keywords:
- vbaxl10.chm168076
ms.prod: excel
api_name:
- Excel.VPageBreaks.Add
ms.assetid: 3196719d-c423-675b-6465-8ac0e9a1c302
ms.date: 06/08/2017
---


# VPageBreaks.Add Method (Excel)

Adds a vertical page break.


## Syntax

 _expression_ . **Add**( **_Before_** )

 _expression_ A variable that represents a **VPageBreaks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Before_|Required| **Object**|A  **[Range](range-object-excel.md)** object. The range to the left of which the new page break will be added.|

### Return Value

A  **[VPageBreak](vpagebreak-object-excel.md)** object that represents the new vertical page break.


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


[VPageBreaks Object](vpagebreaks-object-excel.md)

