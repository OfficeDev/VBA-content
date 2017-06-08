---
title: Application.ProjectBeforeTaskNew Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeTaskNew
ms.assetid: 77418f84-1d82-b227-75f8-c688b7bddf82
ms.date: 06/08/2017
---


# Application.ProjectBeforeTaskNew Event (Project)

Occurs before one or more tasks are created.


## Syntax

 _expression_. **ProjectBeforeTaskNew**( ** _pj_**, ** _Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project in which a task or tasks are being created.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the new task or tasks are not created.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.

The  **ProjectBeforeTaskNew** event doesn't occur when data is merged or appended into a project, during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form.


