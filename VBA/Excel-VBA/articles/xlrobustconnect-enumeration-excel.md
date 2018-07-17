---
title: XlRobustConnect Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlRobustConnect
ms.assetid: 124b8c0f-5120-043e-f226-80d0a7fefe15
ms.date: 06/08/2017
---


# XlRobustConnect Enumeration (Excel)

Specifies how the PivotTable cache or a query table connects to its data source.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlAlways**|1|The PivotTable cache or query table always uses external source information (as defined by the  **SourceConnectionFile** or **SourceDataFile** property) to reconnect.|
| **xlAsRequired**|0|The PivotTable cache or query table uses external source information to reconnect, using the  **Connection** property.|
| **xlNever**|2|The PivotTable cache or query table never uses source information to reconnect.|

