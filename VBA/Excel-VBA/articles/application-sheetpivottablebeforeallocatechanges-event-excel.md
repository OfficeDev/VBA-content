---
title: Application.SheetPivotTableBeforeAllocateChanges Event (Excel)
keywords: vbaxl10.chm504105
f1_keywords:
- vbaxl10.chm504105
ms.prod: excel
api_name:
- Excel.Application.SheetPivotTableBeforeAllocateChanges
ms.assetid: b76cc20d-6251-def7-44d2-504fd6e9cda9
ms.date: 06/08/2017
---


# Application.SheetPivotTableBeforeAllocateChanges Event (Excel)

Occurs before changes are applied to a PivotTable.


## Syntax

 _expression_ . **SheetPivotTableBeforeAllocateChanges**( **_Sh_** , **_TargetPivotTable_** , **_ValueChangeStart_** , **_ValueChangeEnd_** , **_Cancel_** )

 _expression_ A variable that represents a **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The worksheet that contains the PivotTable.|
| _TargetPivotTable_|Required| **[PivotTable](pivottable-object-excel.md)**|The PivotTable that contains the changes to apply.|
| _ValueChangeStart_|Required| **Long**|The index to the first change in the associated  **[PivotTableChangeList](pivottablechangelist-object-excel.md)** collection. The index is specified by the **[Order](valuechange-order-property-excel.md)** property of the **[ValueChange](valuechange-object-excel.md)** object in the **PivotTableChangeList** collection.|
| _ValueChangeEnd_|Required| **Long**|The index to the last change in the associated  **PivotTableChangeList** collection. The index is specified by the **Order** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the changes are not applied to the PivotTable and all edits are lost.|

### Return Value

 **Nothing**


## Remarks

The  **SheetPivotTableBeforeAllocateChanges** event occurs immediately before Excel executes an **UPDATE CUBE** statement to apply all changes to the PivotTable's OLAP data source, and immediately after the user has chosen to apply changes in the user interface.


## See also


#### Concepts


[Application Object](application-object-excel.md)

