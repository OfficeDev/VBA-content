---
title: Workbooks.Add Method (Excel)
keywords: vbaxl10.chm203073
f1_keywords:
- vbaxl10.chm203073
ms.prod: excel
api_name:
- Excel.Workbooks.Add
ms.assetid: ea9f2a2c-3cad-0c35-37b5-82da2f24b876
ms.date: 06/08/2017
---


# Workbooks.Add Method (Excel)

Creates a new workbook. The new workbook becomes the active workbook.


## Syntax

 _expression_ . **Add**( **_Template_** )

 _expression_ A variable that represents a **Workbooks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Template_|Optional| **Variant**|Determines how the new workbook is created. If this argument is a string specifying the name of an existing Microsoft Excel file, the new workbook is created with the specified file as a template. If this argument is a constant, the new workbook contains a single sheet of the specified type. Can be one of the following [XlWBATemplate](xlwbatemplate-enumeration-excel.md) constants: **xlWBATChart** , **xlWBATExcel4IntlMacroSheet** , **xlWBATExcel4MacroSheet** , or **xlWBATWorksheet** . If this argument is omitted, Microsoft Excel creates a new workbook with a number of blank sheets (the number of sheets is set by the **[SheetsInNewWorkbook](application-sheetsinnewworkbook-property-excel.md)** property).|

### Return Value

A  **[Workbook](workbook-object-excel.md)** object that represents the new workbook.


## Remarks

If the  _Template_ argument specifies a file, the file name can include a path.


## See also


#### Concepts


[Workbooks Object](workbooks-object-excel.md)

