---
title: Application.ProjectBeforePrint2 Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforePrint2
ms.assetid: 93e243b7-d765-e3d9-d061-dd98407010d1
ms.date: 06/08/2017
---


# Application.ProjectBeforePrint2 Event (Project)

Occurs before a project is printed. Uses the  **EventInfo** object parameter.


## Syntax

 _expression_. **ProjectBeforePrint2**( ** _pj_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project to be printed.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the project will not be printed.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


