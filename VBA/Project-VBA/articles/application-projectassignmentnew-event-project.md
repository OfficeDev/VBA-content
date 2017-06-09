---
title: Application.ProjectAssignmentNew Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectAssignmentNew
ms.assetid: dcb4acc6-a113-1e93-5f08-e9e68b902b96
ms.date: 06/08/2017
---


# Application.ProjectAssignmentNew Event (Project)

Occurs when a new assignment is created.


## Syntax

 _expression_. **ProjectAssignmentNew**( ** _pj_**, ** _ID_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project where the assignment was just created.|
| _ID_|Required|**Long**|The ID of the assignment that was just created.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


