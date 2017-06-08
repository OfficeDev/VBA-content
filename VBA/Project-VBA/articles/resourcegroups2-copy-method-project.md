---
title: ResourceGroups2.Copy Method (Project)
ms.prod: project-server
api_name:
- Project.ResourceGroups2.Copy
ms.assetid: 3de6fbeb-9067-5ab1-590e-82d2d3c9a136
ms.date: 06/08/2017
---


# ResourceGroups2.Copy Method (Project)

Makes a copy of a group definition for the  **ResourceGroups2** collection and returns a reference to the **[Group2](group2-object-project.md)** object.


## Syntax

 _expression_. **Copy**( ** _Name_**, ** _NewName_** )

 _expression_ An expression that returns a **ResourceGroups2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

### Return Value

 **Group2**


## See also


#### Concepts


[ResourceGroups2 Collection Object](resourcegroups2-object-project.md)

