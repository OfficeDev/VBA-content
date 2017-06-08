---
title: TempVars.Item Property (Access)
keywords: vbaac10.chm14067
f1_keywords:
- vbaac10.chm14067
ms.prod: access
api_name:
- Access.TempVars.Item
ms.assetid: b2b71b6c-cfb4-0b1d-2417-a71725584642
ms.date: 06/08/2017
---


# TempVars.Item Property (Access)

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **TempVar**.


## Syntax

 _expression_. **Item**( ** _Index_** )

 _expression_ A variable that represents a **TempVars** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the  _expression_ argument. If a numeric expression, the _index_ argument must be a number from 0 to the value of the collection's **Count** property minus 1. If a string expression, the _index_ argument must be the name of a member of the collection.|

## See also


#### Concepts


[TempVars Collection](tempvars-object-access.md)

