---
title: Application.WindowGoalAreaChange Event (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowGoalAreaChange
ms.assetid: 1ae33d11-f8aa-e1a2-b59d-9736ce4a6283
ms.date: 06/08/2017
---


# Application.WindowGoalAreaChange Event (Project)

Occurs after a user clicks a different goal area in the Project Guide.


## Syntax

 _expression_. **WindowGoalAreaChange**( ** _Window_**, ** _goalArea_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**|The window where the  **Project Guide** is being changed.|
| _goalArea_|Required|**Long**|The ID of the goal area the user just clicked.|

### Return Value

nothing


## Remarks


 **Note**  The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of a custom Project Guide for new development.

Project events do not occur when the project is embedded in another document or application.


