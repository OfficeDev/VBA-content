---
title: Application.LevelSelected Method (Project)
keywords: vbapj.chm2292
f1_keywords:
- vbapj.chm2292
ms.prod: project-server
api_name:
- Project.Application.LevelSelected
ms.assetid: 1e9383cc-43d3-b479-9b95-cf6fb8cf05b1
ms.date: 06/08/2017
---


# Application.LevelSelected Method (Project)

Levels the selected tasks to resolve resource conflicts or overallocations.


## Syntax

 _expression_. **LevelSelected**( ** _ResolveMethod_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ResolveMethod_|Optional|**PjLevelSelectedOption**|Specifies what to resolve in the leveling process. Can be a  **[PjLevelSelectedOption](pjlevelselectedoption-enumeration-project.md)** constant. The default is **pjResolveSelectedTasks**.|

### Return Value

 **Boolean**


## Remarks

The  **LevelSelected** method corresponds to the **Level Selection** command in the **Level** group on the **Resource** tab. The **Level Selection** command is enabled when more than one task is selected.


