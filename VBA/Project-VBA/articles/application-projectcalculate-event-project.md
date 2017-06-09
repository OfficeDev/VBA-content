---
title: Application.ProjectCalculate Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectCalculate
ms.assetid: 44dbf3f9-4a7d-2e85-aa63-915ea47af008
ms.date: 06/08/2017
---


# Application.ProjectCalculate Event (Project)

Occurs after a project is calculated.


## Syntax

 _expression_. **ProjectCalculate**( ** _pj_**, )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that was calculated.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


