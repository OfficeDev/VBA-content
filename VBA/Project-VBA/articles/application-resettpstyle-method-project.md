---
title: Application.ResetTPStyle Method (Project)
keywords: vbapj.chm1508
f1_keywords:
- vbapj.chm1508
ms.prod: project-server
api_name:
- Project.Application.ResetTPStyle
ms.assetid: aba4187b-5af3-3a8d-7486-038e9bdae0ae
ms.date: 06/08/2017
---


# Application.ResetTPStyle Method (Project)

Resets the specified Team Planner style to the default values.


## Syntax

 _expression_. **ResetTPStyle**( ** _Style_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**PjTeamPlannerStyle**|Can be one of the  **[PjTeamPlannerStyle](pjteamplannerstyle-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

The  **PjTeamPlannerStyle** constants are equivalent to the five styles shown in the **Format** tab of the **Team Planner Tools** in the ribbon, as follows:


|||
|:-----|:-----|
|**Constant**|**Style**|
|**pjTPActualWork**|**Actual Work**|
|**pjTPLateTask**|**Late Task**|
|**pjTPManualTask**|**Manually Scheduled**|
|**pjTPScheduledWork**|**Auto Scheduled**|
|**pjTPSRA**|**External Task**|

## Example

The following line of code resets the border color and fill color of auto-scheduled assignments in the Team Planner to their default values.


```
ResetTPStyle Style:=pjTPScheduledWork
```


