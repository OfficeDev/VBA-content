---
title: PjBoxSet Enumeration (Project)
ms.prod: project-server
api_name:
- Project.PjBoxSet
ms.assetid: 7eea02e0-3bac-cd80-4f19-fc8ce7e1da5c
ms.date: 06/08/2017
---


# PjBoxSet Enumeration (Project)

Contains constants that specify the creation, selection, or movement of a task in the Network Diagram view.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**pjBoxAddToSelection**|0|Selects the task, retaining any existing selection.|
|**pjBoxCreate**|1|Creates a new task, clearing any existing selection.|
|**pjBoxMoveAbsolute**|2|Positions the task relative to the upper left corner of the view. If more than one task is selected and TaskID is not specified, all selected tasks are moved. If TaskID is specified, the selection is cleared and only that task is moved.|
|**pjBoxMoveRelative**|3|Positions the task relative to its current position. If more than one task is selected and TaskID is not specified, all selected tasks are moved. If TaskID is specified, the selection is cleared and only that task is moved.|
|**pjBoxSelect**|4|Selects the task, clearing any existing selection.|
|**pjBoxUnselect**|5|Removes the task from the selection. If more than one task is selected and TaskID is not specified, the box with focus is removed from the selection. If TaskID is specified, only that task is removed from the selection.|

