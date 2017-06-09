---
title: Application.WorkpaneDisplayChange Event (Project)
ms.prod: project-server
api_name:
- Project.Application.WorkpaneDisplayChange
ms.assetid: 8fad51ed-57f5-a34d-6ef6-f699b605c10c
ms.date: 06/08/2017
---


# Application.WorkpaneDisplayChange Event (Project)

Occurs when the Project Guide is hidden or shown.


## Syntax

 _expression_. **WorkpaneDisplayChange**( ** _DisplayState_**, )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DisplayState_|Required|**Boolean**|**True** if the **Project Guide** is shown. **False** if the **Project Guide** is hidden.|

### Return Value

nothing


## Remarks


 **Note**  The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of a custom Project Guide for new development.

Project events do not occur when the project is embedded in another document or application.


