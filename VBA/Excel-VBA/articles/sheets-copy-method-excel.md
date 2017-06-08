---
title: Sheets.Copy Method (Excel)
keywords: vbaxl10.chm152074
f1_keywords:
- vbaxl10.chm152074
ms.prod: excel
api_name:
- Excel.Sheets.Copy
ms.assetid: 8cfee52e-dc0f-a54f-21ba-00a65ba2029c
ms.date: 06/08/2017
---


# Sheets.Copy Method (Excel)

Copies the sheet to another location in the workbook.


## Syntax

 _expression_ . **Copy**( **_Before_** , **_After_** )

 _expression_ A variable that represents a **Sheets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|The sheet before which the copied sheet will be placed. You cannot specify  _Before_ if you specify _After_.|
| _After_|Optional| **Variant**|The sheet after which the copied sheet will be placed. You cannot specify  _After_ if you specify _Before_.|

## Remarks

If you don't specify either  _Before_ or _After_, Microsoft Excel creates a new workbook that contains the copied sheet.


## Example

This example copies Sheet1, placing the copy after Sheet3.


```vb
Worksheets("Sheet1").Copy After:=Worksheets("Sheet3")
```


## See also


#### Concepts


[Sheets Object](sheets-object-excel.md)

