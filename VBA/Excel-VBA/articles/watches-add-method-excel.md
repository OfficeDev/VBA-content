---
title: Watches.Add Method (Excel)
keywords: vbaxl10.chm688073
f1_keywords:
- vbaxl10.chm688073
ms.prod: excel
api_name:
- Excel.Watches.Add
ms.assetid: 18553797-09b9-b99b-c3f3-50864ec2c55a
ms.date: 06/08/2017
---


# Watches.Add Method (Excel)

Adds a range which is tracked when the worksheet is recalculated.


## Syntax

 _expression_ . **Add**( **_Source_** )

 _expression_ A variable that represents a **Watches** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Variant**|The source for the range.|

### Return Value

A  **[Watch](watch-object-excel.md)** object that represents the new range.


## Example

This example creates a summation formula in cell A3 and then adds this cell to the watch facility.


```vb
Sub AddWatch() 
 
 With Application 
 .Range("A1").Formula = 1 
 .Range("A2").Formula = 2 
 .Range("A3").Formula = "=Sum(A1:A2)" 
 .Range("A3").Select 
 .Watches.Add Source:=ActiveCell 
 End With 
 
End Sub
```


## See also


#### Concepts


[Watches Object](watches-object-excel.md)

