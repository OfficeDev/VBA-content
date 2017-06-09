---
title: Workbook.SheetPivotTableBeforeDiscardChanges Event (Excel)
keywords: vbaxl10.chm503105
f1_keywords:
- vbaxl10.chm503105
ms.prod: excel
api_name:
- Excel.Workbook.SheetPivotTableBeforeDiscardChanges
ms.assetid: e8f1ae21-c9ed-6f4d-a85c-d6768060a66f
ms.date: 06/08/2017
---


# Workbook.SheetPivotTableBeforeDiscardChanges Event (Excel)

Occurs before changes to a PivotTable are discarded.


## Syntax

 _expression_ . **SheetPivotTableBeforeDiscardChanges**( **_Sh_** , **_TargetPivotTable_** , **_ValueChangeStart_** , **_ValueChangeEnd_** )

 _expression_ A variable that represents a **[Workbook](workbook-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**||
| _TargetPivotTable_|Required| **[PivotTable](pivottable-object-excel.md)**|The PivotTable that contains the changes to discard.|
| _ValueChangeStart_|Required| **Long**|The index to the first change in the associated  **[PivotTableChangeList](pivottablechangelist-object-excel.md)** object. The index is specified by the **[Order](valuechange-order-property-excel.md)** property of the **[ValueChange](valuechange-object-excel.md)** object in the **PivotTableChangeList** collection.|
| _ValueChangeEnd_|Required| **Long**|The index to the last change in the associated  **PivotTableChangeList** object. The index is specified by the **Order** property of the **ValueChange** object in the **PivotTableChangeList** collection.|

### Return Value

 **Nothing**


## Remarks

Occurs immediately before Excel executes a  **ROLLBACK TRANSACTION** statement against the OLAP data source, if a transaction is still active, and then discards all edited values in the PivotTable, after the user has chosen to discard changes.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

