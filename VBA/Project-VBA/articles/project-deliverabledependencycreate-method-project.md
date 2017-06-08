---
title: Project.DeliverableDependencyCreate Method (Project)
ms.prod: project-server
api_name:
- Project.Project.DeliverableDependencyCreate
ms.assetid: 31ce58fe-3a6a-6151-ebce-b2458728f384
ms.date: 06/08/2017
---


# Project.DeliverableDependencyCreate Method (Project)

Creates a dependency on a deliverable and links the dependency to a task in the project.


## Syntax

 _expression_. **DeliverableDependencyCreate**( ** _DeliverableGuid_**, ** _TaskGuid_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DeliverableGuid_|Required|**String**|The GUID of the deliverable on which to create the dependency.|
| _TaskGuid_|Required|**String**|The GUID of the task to link the dependency.|

### Return Value

 **Boolean**


