---
title: Sheets.Add Method (Excel)
keywords: vbaxl10.chm152073
f1_keywords:
- vbaxl10.chm152073
ms.prod: excel
api_name:
- Excel.Sheets.Add
ms.assetid: db5de750-fd09-2b18-c52b-98d88eeb0ffc
ms.date: 06/08/2017
---


# Sheets.Add Method (Excel)

Creates a new worksheet, chart, or macro sheet. The new worksheet becomes the active sheet.


## Syntax

 _expression_ . **Add**( **_Before_** , **_After_** , **_Count_** , **_Type_** )

 _expression_ A variable that represents a **Sheets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|An object that specifies the sheet before which the new sheet is added.|
| _After_|Optional| **Variant**|An object that specifies the sheet after which the new sheet is added.|
| _Count_|Optional| **Variant**|The number of sheets to be added. The default value is one.|
| _Type_|Optional| **Variant**|Specifies the sheet type. Can be one of the following  **[XlSheetType](xlsheettype-enumeration-excel.md)** constants: **xlWorksheet** , **xlChart** , **xlExcel4MacroSheet** , or **xlExcel4IntlMacroSheet** . If you are inserting a sheet based on an existing template, specify the path to the template. The default value is **xlWorksheet** .|

### Return Value

An Object value that represents the new worksheet, chart, or macro sheet.


## Remarks

If  _Before_ and _After_ are both omitted, the new sheet is inserted before the active sheet.


## Example

This example inserts a new worksheet before the last worksheet in the active workbook.


```vb
ActiveWorkbook.Sheets.Add Before:=Worksheets(Worksheets.Count)
```


## See also


#### Concepts


[Sheets Object](sheets-object-excel.md)

