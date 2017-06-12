---
title: ResourceGroups.Copy Method (Project)
ms.prod: project-server
api_name:
- Project.ResourceGroups.Copy
ms.assetid: 0cf50d60-889b-b599-55be-288aa64f23ee
ms.date: 06/08/2017
---


# ResourceGroups.Copy Method (Project)

Makes a copy of a group definition for the  **ResourceGroups** collection and returns a reference to the **[Group](group-object-project.md)** object.


## Syntax

 _expression_. **Copy**( ** _Name_**, ** _NewName_** )

 _expression_ A variable that represents a **ResourceGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

### Return Value

 **Group**


