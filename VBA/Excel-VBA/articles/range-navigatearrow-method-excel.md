---
title: Range.NavigateArrow Method (Excel)
keywords: vbaxl10.chm144163
f1_keywords:
- vbaxl10.chm144163
ms.prod: excel
api_name:
- Excel.Range.NavigateArrow
ms.assetid: 71e2ce3b-3da8-afd5-7fd3-b922c6f8f1c2
ms.date: 06/08/2017
---


# Range.NavigateArrow Method (Excel)

Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a  **[Range](range-object-excel.md)** object that represents the new selection. This method causes an error if it's applied to a cell without visible tracer arrows.


## Syntax

 _expression_ . **NavigateArrow**( **_TowardPrecedent_** , **_ArrowNumber_** , **_LinkNumber_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TowardPrecedent_|Optional| **Variant**|Specifies the direction to navigate:  **True** to navigate toward precedents, **False** to navigate toward dependent.|
| _ArrowNumber_|Optional| **Variant**|Specifies the arrow number to navigate; corresponds to the numbered reference in the cell's formula.|
| _LinkNumber_|Optional| **Variant**|If the arrow is an external reference arrow, this argument indicates which external reference to follow. If this argument is omitted, the first external reference is followed.|

### Return Value

Variant


## Example

This example navigates along the first tracer arrow from cell A1 on Sheet1 toward the precedent cell. The example should be run on a worksheet containing a formula in cell A1 that includes references to cells D1, D2, and D3 (for example, the formula =D1*D2*D3). Before running the example, display the  **Auditing** toolbar, select cell A1, and click the **Trace Precedents** button.


```vb
Worksheets("Sheet1").Activate 
Range("A1").Select 
ActiveCell.NavigateArrow True, 1
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

