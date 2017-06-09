---
title: PivotTable.RepeatAllLabels Method (Excel)
keywords: vbaxl10.chm235195
f1_keywords:
- vbaxl10.chm235195
ms.prod: excel
api_name:
- Excel.PivotTable.RepeatAllLabels
ms.assetid: 4ca1a7fa-4db6-20da-e37b-37445fee30cf
ms.date: 06/08/2017
---


# PivotTable.RepeatAllLabels Method (Excel)

Specifies whether to repeat item labels for all PivotFields in the specified PivotTable.


## Syntax

 _expression_ . **RepeatAllLabels**( **_Repeat_** )

 _expression_ A variable that represents a **[PivotTable](pivottable-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Repeat_|Required| **[XlPivotFieldRepeatLabels](xlpivotfieldrepeatlabels-enumeration-excel.md)**||

### Return Value

Nothing


## Remarks

Using the  **RepeatAllLabels** method corresponds to the **Repeat All Item Labels** and **Do Not Repeat Item Labels** commands on the **Report Layout** drop-down list of the **PivotTable Tools Design** tab.

To specify whether to repeat item labels for a single PivotField, use the  **[RepeatLabels](pivotfield-repeatlabels-property-excel.md)** property.


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

