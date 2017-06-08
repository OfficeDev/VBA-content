---
title: XlCubeFieldSubType Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlCubeFieldSubType
ms.assetid: 5c5f2390-9bbb-dc46-4aef-5dd47e256c59
ms.date: 06/08/2017
---


# XlCubeFieldSubType Enumeration (Excel)

Specifies the subtype of the CubeField.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlCubeAttribute**|4|Attribute|
| **xlCubeCalculatedMeasure**|5|Calculated Measure|
| **xlCubeHierarchy**|1|Hierarchy|
| **xlCubeImplicitMeasure**|11|An implicit measure|
| **xlCubeKPIGoal**|7|KPI Goal|
| **xlCubeKPIStatus**|8|KPI Status|
| **xlCubeKPITrend**|9|KPI Trend|
| **xlCubeKPIValue**|6|KPI Value|
| **xlCubeKPIWeight**|10|KPI Weight|
| **xlCubeMeasure**|2|Measure|
| **xlCubeSet**|3|Set|

## Remarks


 **Note**  The values have ?Cube? in the name in order to not overlap with the  **xlMeasure** and **xlSet** values for the **CubeFieldType** property. If the names are the same, autocomplete will not work in the Visual Basic environment because it finds ambiguous values.


