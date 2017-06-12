---
title: Application.TaskMoveToStatusDate Method (Project)
keywords: vbapj.chm2290
f1_keywords:
- vbapj.chm2290
ms.prod: project-server
api_name:
- Project.Application.TaskMoveToStatusDate
ms.assetid: 100ec970-ca52-2ac8-f367-c346c40e4c61
ms.date: 06/08/2017
---


# Application.TaskMoveToStatusDate Method (Project)

Moves completed or incomplete parts of one or more selected tasks to the status date. 


## Syntax

 _expression_. **TaskMoveToStatusDate**( ** _MoveCompleted_**, ** _MoveIncomplete_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MoveCompleted_|Optional|**Boolean**|**True** if the completed parts of tasks are moved to the status date; otherwise, **False**. The default is **False**.|
| _MoveIncomplete_|Optional|**Boolean**|**True** if the incomple parts of tasks are moved to the status date; otherwise, **False**. The default is **True**.|

### Return Value

 **Boolean**


## Remarks

To set or change the status date, click  **Project Information** on the Project tab on the Ribbon. The Project Information dialog box includes the **Status date** field. If the status date value is "NA", no status date is set. In that case, the current date is the status date.

If both the  _MoveCompleted_ and _MoveIncomplete_ arguments are **False**, **TaskMoveToStatusDate** takes no action but still returns **True**. If both arguments are **True**, **TaskMoveToStatusDate** moves only the incomplete parts to the status date.

The  **TaskMoveToStatusDate** method corresponds to the **Incomplete Parts to Status Date** or **Completed Parts to Status Date** commands in the **Move Task** drop-down menu on the **TASK** ribbon. The **[TaskMove](application-taskmove-method-project.md)** method corresponds to other commands on the **Move Task** drop-down menu.


