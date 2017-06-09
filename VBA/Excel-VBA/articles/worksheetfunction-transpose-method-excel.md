---
title: WorksheetFunction.Transpose Method (Excel)
keywords: vbaxl10.chm137117
f1_keywords:
- vbaxl10.chm137117
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Transpose
ms.assetid: 327aaf19-c226-5251-9bec-eadc4546d53a
ms.date: 06/08/2017
---


# WorksheetFunction.Transpose Method (Excel)

Returns a vertical range of cells as a horizontal range, or vice versa. TRANSPOSE must be entered as an array formula in a range that has the same number of rows and columns, respectively, as an array has columns and rows. Use TRANSPOSE to shift the vertical and horizontal orientation of an array on a worksheet.


## Syntax

 _expression_ . **Transpose**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - an array or range of cells on a worksheet that you want to transpose. The transpose of an array is created by using the first row of the array as the first column of the new array, the second row of the array as the second column of the new array, and so on.|

### Return Value

Variant


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

