---
title: Application.ProjectBeforeClose2 Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeClose2
ms.assetid: 24b43d85-f99c-915c-47fe-0df5875fc479
ms.date: 06/08/2017
---


# Application.ProjectBeforeClose2 Event (Project)

Occurs before a project is closed. Uses the  **EventInfo** object parameter.


## Syntax

 _expression_. **ProjectBeforeClose2**( ** _pj_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project to be closed|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the project will not be closed.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


