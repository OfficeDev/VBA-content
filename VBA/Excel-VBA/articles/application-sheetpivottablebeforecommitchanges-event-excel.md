---
title: Application.SheetPivotTableBeforeCommitChanges Event (Excel)
keywords: vbaxl10.chm504106
f1_keywords:
- vbaxl10.chm504106
ms.prod: excel
api_name:
- Excel.Application.SheetPivotTableBeforeCommitChanges
ms.assetid: ba586d2e-772a-24e3-0886-fb309f17ebf6
ms.date: 06/08/2017
---


# Application.SheetPivotTableBeforeCommitChanges Event (Excel)

Occurs before changes are committed against the OLAP data source for a PivotTable.


## Syntax

 _expression_ . **SheetPivotTableBeforeCommitChanges**( **_Sh_** , **_TargetPivotTable_** , **_ValueChangeStart_** , **_ValueChangeEnd_** , **_Cancel_** )

 _expression_ A variable that represents a **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The worksheet that contains the PivotTable.|
| _TargetPivotTable_|Required| **[PivotTable](pivottable-object-excel.md)**|The PivotTable that contains the changes to commit.|
| _ValueChangeStart_|Required| **Long**|The index to the first change in the associated  **[PivotTableChangeList](pivottablechangelist-object-excel.md)** object. The index is specified by the **[Order](valuechange-order-property-excel.md)** property of the **[ValueChange](valuechange-object-excel.md)** object in the **PivotTableChangeList** collection.|
| _ValueChangeEnd_|Required| **Long**|The index to the last change in the associated  **PivotTableChangeList** object. The index is specified by the **Order** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the changes are not committed against the OLAP data source of the PivotTable.|

### Return Value

 **Nothing**


## Remarks

The  **SheetPivotTableBeforeCommitChanges** event occurs immediately before Excel executes a **COMMIT TRANSACTION** against the PivotTable's OLAP data source, and immediately after the user has chosen to save changes for the whole PivotTable.


## See also


#### Concepts


[Application Object](application-object-excel.md)

