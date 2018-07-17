---
title: TaskDependencies.Add Method (Project)
ms.prod: project-server
api_name:
- Project.TaskDependencies.Add
ms.assetid: 37e67ab2-ca7b-26c2-50e7-8a933b746489
ms.date: 06/08/2017
---


# TaskDependencies.Add Method (Project)

Adds a  **TaskDependency** object to a **TaskDependencies** collection.


## Syntax

 _expression_. **Add**( ** _From_**, ** _Type_**, ** _Lag_** )

 _expression_ A variable that represents a **TaskDependencies** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _From_|Required|**Object**|The  **Task** object specified becomes a predecessor of the task specified by expression.|
| _Type_|Optional|**Long**|The type of relationship between the linked tasks. Can be one of the  **[PjTaskLinkType](pjtasklinktype-enumeration-project.md)** constants. The default value is **pjFinishToStart**.|
| _Lag_|Optional|**Variant**|The duration of lag time between linked tasks. To specify lead time between tasks, use a negative value. String values default to days unless otherwise specified. Non-string values are interpreted as minutes. The default value is 0.|

### Return Value

 **TaskDependency**


## See also


#### Concepts


[TaskDependencies Collection Object](taskdependencies-object-project.md)
