---
title: Application.GanttRollup Method (Project)
keywords: vbapj.chm2119
f1_keywords:
- vbapj.chm2119
ms.prod: project-server
api_name:
- Project.Application.GanttRollup
ms.assetid: 8bb5ef38-d0c7-7425-a6ac-e50c7ae979d8
ms.date: 06/08/2017
---


# Application.GanttRollup Method (Project)

Specifies the rollup behavior of bars on the Gantt Chart.


## Syntax

 _expression_. **GanttRollup**( ** _AlwaysRollup_**, ** _HideWhenSummaryExpanded_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AlwaysRollup_|Optional|**Boolean**|**True** if rolled-up Gantt bars always display. The default value is **False**.|
| _HideWhenSummaryExpanded_|Optional|**Boolean**|**True** if rolled-up Gantt bars should be hidden when summary tasks are expanded. The default value is **False**.|

### Return Value

 **Boolean**


