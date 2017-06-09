---
title: PivotTable.PivotSelect Method (Excel)
keywords: vbaxl10.chm235137
f1_keywords:
- vbaxl10.chm235137
ms.prod: excel
api_name:
- Excel.PivotTable.PivotSelect
ms.assetid: e9beda74-c022-3ba7-b3af-d607024846f2
ms.date: 06/08/2017
---


# PivotTable.PivotSelect Method (Excel)

Selects part of a PivotTable report.


## Syntax

 _expression_ . **PivotSelect**( **_Name_** , **_Mode_** , **_UseStandardName_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The part of the PivotTable report to select.|
| _Mode_|Optional| **[XlPTSelectionMode](xlptselectionmode-enumeration-excel.md)**|Specifies the structured selection mode.|
| _UseStandardName_|Optional| **Variant**| **True** for recorded macros that will play back in other locales.|

## Remarks

You can use the specified mode only to select the corresponding item in the PivotTable report. For example, you cannot select data and labels by using  **xlButton** mode; likewise, you cannot select buttons by using **xlDataOnly** mode.


## Example

This example selects all date labels in the first PivotTable report on worksheet one.


```vb
Worksheets(1).PivotTables(1).PivotSelect "date[All]", xlLabelOnly
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

