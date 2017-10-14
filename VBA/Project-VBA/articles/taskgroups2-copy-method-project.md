---
title: TaskGroups2.Copy Method (Project)
ms.prod: project-server
api_name:
- Project.TaskGroups2.Copy
ms.assetid: 7afc3518-e5bb-52be-0a45-edb436381250
ms.date: 06/08/2017
---


# TaskGroups2.Copy Method (Project)

Makes a copy of a group definition for the  **TaskGroups2** collection and returns a reference to the **[Group2](group2-object-project.md)** object.


## Syntax

 _expression_. **Copy**( ** _Name_**, ** _NewName_** )

 _expression_ An expression that returns a **TaskGroups2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

### Return Value

 **Group2**


## See also


#### Concepts


[TaskGroups2 Collection Object](taskgroups2-object-project.md)

