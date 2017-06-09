---
title: PivotField.IncludeNewItemsInFilter Property (Excel)
keywords: vbaxl10.chm240145
f1_keywords:
- vbaxl10.chm240145
ms.prod: excel
api_name:
- Excel.PivotField.IncludeNewItemsInFilter
ms.assetid: af77670a-a7ae-4797-6afb-1be17e611a99
ms.date: 06/08/2017
---


# PivotField.IncludeNewItemsInFilter Property (Excel)

Allows developers to specify whether excluded or included items should be tracked when manual filtering is applied to the PivotField. Read/write  **Boolean** .


## Syntax

 _expression_ . **IncludeNewItemsInFilter**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

The default value for this property is  **False** .

When manual filtering is applied, developers can set the  **IncludeNewItemsInFilter** property to **True** to track excluded items. They can set the property to **False** to track included items. If new items are added to the source data after the **IncludeNewItemsInFilter** property is set to **True** , the new items appear in the PivotTable after the next refresh operation because they are not in the collection of items that Excel is tracking.

For OLAP hierarchies this setting is set on the  **CubeField** object, and trying to set it on PivotFields that are part of a hierarchy will fail (with a run-time error). You can get this setting from the PivotFields that are part of a hierarchy and it will return the same as the corresponding **CubeField.IncludeNewItemsInFilter** property.

Toggling this setting will clear the following collections:  **HiddenItemsList** , **HiddenItems** , **VisibleItemsList** , and **VisibleItems** . When **IncludeNewItemsInFilter** is set to **False** , the **HiddenItemsList** and **HiddenItems** collections are empty and items cannot be added to them. A run-time error is returned when attempting to add items. When **IncludeNewItemsInFilter** is set to **True** , the **VisibleItemsList** and **VisibleItems** collections are empty and items cannot be added to them. A run-time error is returned when attempting to add items.

In PivotTables, this setting is set on the  **PivotField** object. Toggling this setting will not change the filter state.


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

