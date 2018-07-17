---
title: Groups.Copy Method (Project)
ms.prod: project-server
api_name:
- Project.Groups.Copy
ms.assetid: fa53fb17-be05-ab03-c08b-a2c9034b7da6
ms.date: 06/08/2017
---


# Groups.Copy Method (Project)

Makes a copy of a group definition for the  **Groups** collection and returns a reference to the **[Group](group-object-project.md)** object.


## Syntax

 _expression_. **Copy**( ** _Name_**, ** _NewName_** )

 _expression_ A variable that represents a **Groups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

### Return Value

 **Group**


## See also


#### Concepts


[Groups Collection Object](groups-object-project.md)
