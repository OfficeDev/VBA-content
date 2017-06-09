---
title: SlicerCache.WorkbookConnection Property (Excel)
keywords: vbaxl10.chm897076
f1_keywords:
- vbaxl10.chm897076
ms.prod: excel
api_name:
- Excel.SlicerCache.WorkbookConnection
ms.assetid: ffe4fcbc-025e-6349-aaee-39a938b61e1e
ms.date: 06/08/2017
---


# SlicerCache.WorkbookConnection Property (Excel)

Gets or sets the  **[WorkbookConnection](workbookconnection-object-excel.md)** object that represents the data connection used by the specified slicer. Read/Write.


## Syntax

 _expression_ . **WorkbookConnection**

 _expression_ A variable that represents a **[SlicerCache](slicercache-object-excel.md)** object.


### Return Value

 **WorkbookConnection**


## Remarks

The  **WorkbookConnection** property only applies to slicers that are based on external data sources ( **SlicerCache** . **SourceType** = **xlExternal** ). Attempting to access the **WorkbookConnection** property for slicers that are connected to PivotTables based on workbook ranges or lists ( **SlicerCache** . **SourceType** = **xlDatabase** ) generates a run-time error.

The workbook connection value must be unique. Setting the workbook connection to a value that already exists generates a run-time error. 


## See also


#### Concepts


[SlicerCache Object](slicercache-object-excel.md)

