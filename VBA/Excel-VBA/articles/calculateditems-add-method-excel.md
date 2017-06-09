---
title: CalculatedItems.Add Method (Excel)
keywords: vbaxl10.chm250078
f1_keywords:
- vbaxl10.chm250078
ms.prod: excel
api_name:
- Excel.CalculatedItems.Add
ms.assetid: 2a7dff2b-c874-1579-1e95-78841a91e6cd
ms.date: 06/08/2017
---


# CalculatedItems.Add Method (Excel)

Creates a new calculated item. Returns a  **[PivotItem](pivotitem-object-excel.md)** object.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Formula_** , **_UseStandardFormula_** )

 _expression_ A variable that represents a **CalculatedItems** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the item.|
| _Formula_|Required| **String**|The formula for the item.|
| _UseStandardFormula_|Optional| **Variant**| **False** (default) for upward compatibility. **True** for strings contained in any arguments that are item names, will be interpreted as having been formatted in standard U.S. English instead of local settings.|

### Return Value

A  **PivotItem** object that represents the new calculated item.


## See also


#### Concepts


[CalculatedItems Collection](calculateditems-object-excel.md)

