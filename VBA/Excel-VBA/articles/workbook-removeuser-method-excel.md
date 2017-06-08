---
title: Workbook.RemoveUser Method (Excel)
keywords: vbaxl10.chm199138
f1_keywords:
- vbaxl10.chm199138
ms.prod: excel
api_name:
- Excel.Workbook.RemoveUser
ms.assetid: f0a978a0-7bcf-3af4-a01a-831c6c854989
ms.date: 06/08/2017
---


# Workbook.RemoveUser Method (Excel)

Disconnects the specified user from the shared workbook.


## Syntax

 _expression_ . **RemoveUser**( **_Index_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The user index.|

## Example

This example disconnects user two from the shared workbook.


```vb
Workbooks(2).RemoveUser 2
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

