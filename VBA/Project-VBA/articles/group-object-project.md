---
title: Group Object (Project)
ms.prod: project-server
api_name:
- Project.Group
ms.assetid: e3756818-f051-1ae4-5402-0398e568ebfc
ms.date: 06/08/2017
---


# Group Object (Project)

Represents a group definition. A  **Group** object is a member of the **[ResourceGroups](resourcegroups-object-project.md)** collection or the **[TaskGroups](taskgroups-object-project.md)** collection.
 


## Remarks

 **Using the Group Object**
 

 
Use  `TaskGroups(Index)` or `ResourceGroups(Index)`, where *Index* is the group definition index or group definition name, to return a **Group** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```
ActiveProject.ResourceGroups("Standard Rate").ShowSummary = True
```


## Methods



|**Name**|
|:-----|
|[Delete](group-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](group-application-property-project.md)|
|[GroupAssignments](group-groupassignments-property-project.md)|
|[GroupCriteria](group-groupcriteria-property-project.md)|
|[Index](group-index-property-project.md)|
|[Name](group-name-property-project.md)|
|[Parent](group-parent-property-project.md)|
|[ShowSummary](group-showsummary-property-project.md)|

