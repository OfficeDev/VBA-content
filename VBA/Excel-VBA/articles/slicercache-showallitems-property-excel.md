---
title: SlicerCache.ShowAllItems Property (Excel)
keywords: vbaxl10.chm897088
f1_keywords:
- vbaxl10.chm897088
ms.prod: excel
api_name:
- Excel.SlicerCache.ShowAllItems
ms.assetid: 72622510-b644-db1b-2905-4eaba53b0ecb
ms.date: 06/08/2017
---


# SlicerCache.ShowAllItems Property (Excel)

Returns or sets whether slicers connected to the specified slicer cache display items that have been deleted from in the corresponding PivotCache. Read/write


## Syntax

 _expression_ . **ShowAllItems**

 _expression_ A variable that represents a **[SlicerCache](slicercache-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

When the  **ShowAllItems** property is set to **True** (the default), items that have been deleted from the source data are displayed in the slicers connected to the specified slicer cache. The **ShowAllItems** property corresponds to the setting of the **Show items deleted from the data source** check box in the **Slicer Settings** dialog box.

The  **ShowAllItems** property applies only to slicers that are based on workbook ranges or lists ( **SlicerCache** . **SourceType** = **xlDatabase** ), or to slicers that are based on relational data sources ( **SlicerCache** . **SourceType** = **xlExternal** and **SlicerCache** . **[OLAP](slicercache-olap-property-excel.md)** = **False** ). Attempting to set the **ShowAllItems** property for slicers that are connected to PivotTables based on external OLAP data sources ( **SlicerCache** . **OLAP** = **True** ) generates a run-time error.


## See also


#### Concepts


[SlicerCache Object](slicercache-object-excel.md)

