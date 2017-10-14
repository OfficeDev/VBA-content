---
title: Application.NewTasksStartOn Method (Project)
keywords: vbapj.chm2295
f1_keywords:
- vbapj.chm2295
ms.prod: project-server
api_name:
- Project.Application.NewTasksStartOn
ms.assetid: c5009674-105e-a861-56f0-4847926d6c36
ms.date: 06/08/2017
---


# Application.NewTasksStartOn Method (Project)

Specifies how the start date of a new task is set.


## Syntax

 _expression_. **NewTasksStartOn**( ** _StartOnDate_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StartOnDate_|Optional|**PjNewTasksStartOnDate**|Specifies whether new tasks start on the project date, the current date, or no date. Can be one of the  **[PjNewTasksStartOnDate](pjnewtasksstartondate-enumeration-project.md)** constants. The default is **pjProjectDate**.|

### Return Value

 **Boolean**


## Remarks

The  **NewTasksStartOn** method corresponds to the **New tasks created** setting on the **Schedule** tab of the **Project Options** dialog box.


