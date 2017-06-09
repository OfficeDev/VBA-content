---
title: Group2 Object (Project)
ms.prod: project-server
api_name:
- Project.Group2
ms.assetid: a7a61fa4-e752-006e-a47e-03987b04f01c
ms.date: 06/08/2017
---


# Group2 Object (Project)

Represents a group definition where the group hierarchy can be maintained. A  **Group2** object is a member of a **[Groups2](groups2-object-project.md)**, **[ResourceGroups2](resourcegroups2-object-project.md)**, or **[TaskGroups2](taskgroups2-object-project.md)** collection.
 


## Remarks

The  **Group2** object includes the **[MaintainHierarchy](group2-maintainhierarchy-property-project.md)** property.
 

 
 **Using the Group Object**
 

 
Use  `TaskGroups2(Index)` or `ResourceGroups2(Index)`, where *Index* is the group definition index or group definition name, to return a **Group2** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```
ActiveProject.ResourceGroups2("Standard Rate").ShowSummary = True
```


## Methods



|**Name**|
|:-----|
|[Delete](group2-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](group2-application-property-project.md)|
|[GroupAssignments](group2-groupassignments-property-project.md)|
|[GroupCriteria](group2-groupcriteria-property-project.md)|
|[Index](group2-index-property-project.md)|
|[MaintainHierarchy](group2-maintainhierarchy-property-project.md)|
|[Name](group2-name-property-project.md)|
|[Parent](group2-parent-property-project.md)|
|[ShowSummary](group2-showsummary-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
