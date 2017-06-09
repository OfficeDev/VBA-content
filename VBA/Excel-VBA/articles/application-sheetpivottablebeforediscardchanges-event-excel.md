---
title: Application.SheetPivotTableBeforeDiscardChanges Event (Excel)
keywords: vbaxl10.chm504107
f1_keywords:
- vbaxl10.chm504107
ms.prod: excel
api_name:
- Excel.Application.SheetPivotTableBeforeDiscardChanges
ms.assetid: 8623adc6-d256-bebb-fe35-8710390af19f
ms.date: 06/08/2017
---


# Application.SheetPivotTableBeforeDiscardChanges Event (Excel)

Occurs before changes to a PivotTable are discarded.


## Syntax

 _expression_ . **SheetPivotTableBeforeDiscardChanges**( **_Sh_** , **_TargetPivotTable_** , **_ValueChangeStart_** , **_ValueChangeEnd_** )

 _expression_ A variable that represents a **[Application](application-object-excel.md)** object.


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


[Application Object](application-object-excel.md)

