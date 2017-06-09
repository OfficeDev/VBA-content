---
title: Application.ProjectBeforeClose Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeClose
ms.assetid: 90e75c72-03f9-25ab-1339-94d9ff8933a2
ms.date: 06/08/2017
---


# Application.ProjectBeforeClose Event (Project)

Occurs before a project is closed.


## Syntax

 _expression_. **ProjectBeforeClose**( ** _pj_**, ** _Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project to be closed|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the project will not be closed.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


