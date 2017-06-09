---
title: Application.GoalAreaHighlight Method (Project)
keywords: vbapj.chm131221
f1_keywords:
- vbapj.chm131221
ms.prod: project-server
api_name:
- Project.Application.GoalAreaHighlight
ms.assetid: 56146d8b-f986-0ba7-3661-26b508db3ec8
ms.date: 06/08/2017
---


# Application.GoalAreaHighlight Method (Project)

Highlights a goal area on the  **Project Guide** toolbar to indicate it is currently selected. Deprecated in _pjgenericshort_.


## Syntax

 _expression_. **GoalAreaHighlight**( ** _goalArea_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _goalArea_|Required|**Long**|The ID of the goal area to highlight. For example, setting the  _goalArea_ argument to 1 highlights the first goal area in the **Goal Bar**.|

## Remarks


 **Note**  The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of the Project Guide for new development.


