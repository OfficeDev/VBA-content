---
title: Project.BeforePrint Event (Project)
ms.prod: project-server
api_name:
- Project.Project.BeforePrint
ms.assetid: df66b52b-4c7b-e3e1-d8ff-66416edcb378
ms.date: 06/08/2017
---


# Project.BeforePrint Event (Project)

Occurs before a project is printed.


## Syntax

 _expression_. **BeforePrint**( ** _pj_**, )

 _expression_ An expression that returns a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that will be printed.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application. 


