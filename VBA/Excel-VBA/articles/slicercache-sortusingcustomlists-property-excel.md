---
title: SlicerCache.SortUsingCustomLists Property (Excel)
keywords: vbaxl10.chm897087
f1_keywords:
- vbaxl10.chm897087
ms.prod: excel
api_name:
- Excel.SlicerCache.SortUsingCustomLists
ms.assetid: 61c156fe-67cf-f6e8-4fce-bc617c9a1e03
ms.date: 06/08/2017
---


# SlicerCache.SortUsingCustomLists Property (Excel)

Returns or sets whether items in the specified slicer cache will be sorted by the custom lists. Read/write


## Syntax

 _expression_ . **SortUsingCustomLists**

 _expression_ A variable that represents a **[SlicerCache](slicercache-object-excel.md)** object.


## Remarks

The  **SortUsingCustomLists** property corresponds to the setting of the **Use Custom Lists when sorting check box** of the **Slicer Settings** dialog box. To access the custom lists associated with the current installation of Excel, click the **File** tab, click **Options**, click  **Advanced**, and then click  **Edit Custom Lists** under the **General** category.

The  **SortUsingCustomLists** property only applies to slicers that are filtering non-OLAP data sources. Attempting to access this property from a slicer cache that is filtering an OLAP data source ( **SlicerCache** . **[OLAP](slicercache-olap-property-excel.md)** = **True** ) generates a run-time error.


## See also


#### Concepts


[SlicerCache Object](slicercache-object-excel.md)

