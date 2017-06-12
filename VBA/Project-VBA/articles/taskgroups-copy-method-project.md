---
title: TaskGroups.Copy Method (Project)
ms.prod: project-server
api_name:
- Project.TaskGroups.Copy
ms.assetid: e69fe06d-3855-a8ac-32fe-752ff280fe85
ms.date: 06/08/2017
---


# TaskGroups.Copy Method (Project)

Makes a copy of a group definition for the  **TaskGroups** collection and returns a reference to the **[Group](group-object-project.md)** object.


## Syntax

 _expression_. **Copy**( ** _Name_**, ** _NewName_** )

 _expression_ A variable that represents a **TaskGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

### Return Value

 **Group**


