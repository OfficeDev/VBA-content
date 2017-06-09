---
title: CubeField.IncludeNewItemsInFilter Property (Excel)
keywords: vbaxl10.chm668095
f1_keywords:
- vbaxl10.chm668095
ms.prod: excel
api_name:
- Excel.CubeField.IncludeNewItemsInFilter
ms.assetid: 7c9ccb66-5a8c-ced0-c024-2336e85f00db
ms.date: 06/08/2017
---


# CubeField.IncludeNewItemsInFilter Property (Excel)

The  **IncludeNewItemsInFilter** property is used to track included/excluded items in OLAP PivotTables. Read/write.


## Syntax

 _expression_ . **IncludeNewItemsInFilter**

 _expression_ A variable that represents a **CubeField** object.


## Remarks

Default value is  **False** .

When this setting is set to  **True** , excluded items are tracked when manual filtering is applied. When this setting is set to **False** , included items are tracked when manual filtering is applied.

When  **IncludeNewItemsInFilter** is set to **False** , the **HiddenItemsList** and **HiddenItems** collections are empty and items cannot be added to them.

When  **IncludeNewItemsInFilter** is set to **True** , the **VisibleItemsList** and **VisibleItems** collections are empty and items cannot be added to them.


## See also


#### Concepts


[CubeField Object](cubefield-object-excel.md)

