---
title: PivotField.RepeatLabels Property (Excel)
keywords: vbaxl10.chm240160
f1_keywords:
- vbaxl10.chm240160
ms.prod: excel
api_name:
- Excel.PivotField.RepeatLabels
ms.assetid: abc7e5f7-4633-38b3-d5a8-41bfa463077d
ms.date: 06/08/2017
---


# PivotField.RepeatLabels Property (Excel)

Returns or sets whether item labels are repeated in the PivotTable for the specified PivotField. Read/write


## Syntax

 _expression_ . **RepeatLabels**

 _expression_ A variable that represents a **[PivotField](pivotfield-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

 **True** if item labels are repeated for the specified PivotField; otherwise **False** .

The setting of the  **RepeatLabels** property corresponds to the **Repeat item labels** check box on the ** Layout &; Print** tab of the **Field Settings** dialog box for a field in a PivotTable.

To specify whether to repeat item labels for all PivotFields in a PivotTable in a single operation, use the  **[RepeatAllLabels](pivottable-repeatalllabels-method-excel.md)** method.


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

