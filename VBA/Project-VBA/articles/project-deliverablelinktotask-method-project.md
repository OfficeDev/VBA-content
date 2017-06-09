---
title: Project.DeliverableLinkToTask Method (Project)
ms.prod: project-server
api_name:
- Project.Project.DeliverableLinkToTask
ms.assetid: b3cfea3d-dc49-52a7-2e10-3d1f12cefbc1
ms.date: 06/08/2017
---


# Project.DeliverableLinkToTask Method (Project)

Links a deliverable to a task.


## Syntax

 _expression_. **DeliverableLinkToTask**( ** _DeliverableGuid_**, ** _TaskGuid_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DeliverableGuid_|Required|**String**|The GUID of the deliverable to link.|
| _TaskGuid_|Required|**String**|The GUID of the task to link.|

### Return Value

 **Boolean**


