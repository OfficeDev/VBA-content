---
title: Project.DeliverableCreate Method (Project)
ms.prod: project-server
api_name:
- Project.Project.DeliverableCreate
ms.assetid: 538f8143-0c0d-b9fa-9219-5405f4bd5046
ms.date: 06/08/2017
---


# Project.DeliverableCreate Method (Project)

Creates a deliverable for a published project that has a project workspace.


## Syntax

 _expression_. **DeliverableCreate**( ** _DeliverableName_**, ** _DeliverableStartDate_**, ** _DeliverableFinishDate_**, ** _TaskGuid_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DeliverableName_|Required|**String**|Name of the deliverable.|
| _DeliverableStartDate_|Required|**Variant**|Start date for the deliverable.|
| _DeliverableFinishDate_|Required|**Variant**|Finish date for the deliverable.|
| _TaskGuid_|Required|**String**|GUID of the task to which to link the deliverable.|

### Return Value

 **String**


