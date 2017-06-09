---
title: XlPivotTableVersionList Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlPivotTableVersionList
ms.assetid: a9b1ea64-53a1-0fd5-208e-b609b31c1c64
ms.date: 06/08/2017
---


# XlPivotTableVersionList Enumeration (Excel)

Specifies the version of a PivotTable or a PivotCache. Creating PivotTables with a specific version ensures that tables created in Excel behave in the same manner as they did in the corresponding version of Excel.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlPivotTableVersion2000**|0|Excel 2000|
| **xlPivotTableVersion10**|1|Excel 2002|
| **xlPivotTableVersion11**|2|Excel 2003|
| **xlPivotTableVersion12**|3|Excel 2007|
| **xlPivotTableVersion14**|4|Excel 2010|
| **xlPivotTableVersion15**|5|Excel 2013|
| **xlPivotTableVersionCurrent**|-1|Provided only for backward compatibility|

## Remarks


 **Note**   _xlPivotTableVersionCurrent_ is included only for backward compatibility reasons. It cannot be used with new **PivotCache** and **PivotTable** objects. There are no differences in behavior between _xlPivotTableVersion11_ and _xlPivotTableVersion10_ .


