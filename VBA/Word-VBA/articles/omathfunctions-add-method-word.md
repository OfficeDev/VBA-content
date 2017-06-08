---
title: OMathFunctions.Add Method (Word)
keywords: vbawd10.chm44302440
f1_keywords:
- vbawd10.chm44302440
ms.prod: word
api_name:
- Word.OMathFunctions.Add
ms.assetid: 2292e297-6d24-cd73-971b-146be1edcb0a
ms.date: 06/08/2017
---


# OMathFunctions.Add Method (Word)

Inserts a new structure, such as a fraction, into an equation at the specified position and returns an  **OMathFunction** object that represents the structure.


## Syntax

 _expression_ . **Add**( **_Range_** , **_Type_** , **_NumArgs_** , **_NumCols_** )

 _expression_ An expression that returns a **OMathFunctions** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**| The place at which to insrt an equation.|
| _Type_|Required| **WdOMathFunctionType**|The type of equation to insert.|
| _NumArgs_|Optional| **Variant**| The number of arguments in the equation.|
| _NumCols_|Optional| **Variant**|The number of columns in the equation.|

### Return Value

OMathFunction


## See also


#### Concepts


[OMathFunctions Collection](omathfunctions-object-word.md)

