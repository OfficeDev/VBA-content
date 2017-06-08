---
title: Range.InsertIndent Method (Excel)
keywords: vbaxl10.chm144148
f1_keywords:
- vbaxl10.chm144148
ms.prod: excel
api_name:
- Excel.Range.InsertIndent
ms.assetid: 1e004333-a64e-55e4-cf8a-d15e47236f94
ms.date: 06/08/2017
---


# Range.InsertIndent Method (Excel)

Adds an indent to the specified range.


## Syntax

 _expression_ . **InsertIndent**( **_InsertAmount_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _InsertAmount_|Required| **Long**|The amount to be added to the current indent.|

## Remarks

Using this method to set the indent level to a number less than 0 (zero) or greater than 15 causes an error.

Use the  **IndentLevel** property to return the indent level for a range.


## Example

This example decreases the indent level in cell A10.


```vb
With Range("a10") 
 .InsertIndent -1 
End With
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

