---
title: PivotCache.CommandText Property (Excel)
keywords: vbaxl10.chm227087
f1_keywords:
- vbaxl10.chm227087
ms.prod: excel
api_name:
- Excel.PivotCache.CommandText
ms.assetid: 07921bda-74fe-2a41-15f7-16068ce49a31
ms.date: 06/08/2017
---


# PivotCache.CommandText Property (Excel)

Returns or sets the command string for the specified data source. Read/write  **Variant** .


## Syntax

 _expression_ . **CommandText**

 _expression_ An expression that returns a **PivotCache** object.


## Remarks

For OLE DB sources, the  **[CommandType](pivotcache-commandtype-property-excel.md)** property describes the value of the **CommandText** property.

For ODBC sources, setting the  **CommandText** causes the data to be refreshed.


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

