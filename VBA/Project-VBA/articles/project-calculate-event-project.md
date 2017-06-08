---
title: Project.Calculate Event (Project)
ms.prod: project-server
api_name:
- Project.Project.Calculate
ms.assetid: cba7feb3-c0e4-96ec-d2fa-eaccfa640c5a
ms.date: 06/08/2017
---


# Project.Calculate Event (Project)

Occurs when a project schedule is recalculated.


## Syntax

 _expression_. **Calculate**( ** _pj_**, )

 _expression_ An expression that returns a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that is rescheduled.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application. 


