---
title: Application.ProjectBeforePrint Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforePrint
ms.assetid: 7cc8de23-c3e3-81df-ae26-37c4e639dd81
ms.date: 06/08/2017
---


# Application.ProjectBeforePrint Event (Project)

Occurs before a project is printed.


## Syntax

 _expression_. **ProjectBeforePrint**( ** _pj_**, ** _Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**| The project to be printed.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the project will not be printed.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


