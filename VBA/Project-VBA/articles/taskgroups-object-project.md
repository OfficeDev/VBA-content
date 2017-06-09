---
title: TaskGroups Object (Project)
ms.prod: project-server
ms.assetid: 76d01102-cc38-36c1-f2fb-c5155f3056db
ms.date: 06/08/2017
---


# TaskGroups Object (Project)

Represents all the task-based group definitions.  **TaskGroups** is a collection of **[Group](group-object-project.md)** objects.
 


## Remarks

For task groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, use the  **[TaskGroups2](taskgroups2-object-project.md)** collection object.
 

 

## Example

 **Using the TaskGroups Collection**
 

 
Use the  **[TaskGroups](project-taskgroups-property-project.md)** property to return a **TaskGroups** collection. The following example lists the names of all the task groups in the active project.
 

 



```
Dim tg As Group 
Dim tGroups As String 
 
For Each tg in ActiveProject.TaskGroups 
 tGroups = tGroups &amp; tg.Name &amp; vbCrLf 
Next tg 
 
MsgBox tGroups
```

Use the  **[Add](taskgroups-add-method-project.md)** method to add a **Group** object to the **TaskGroups** collection. The following example creates a new group that groups tasks by whether they are overallocated and then modifies the criterion so that overallocated tasks are sorted in descending order.
 

 



```
ActiveProject.TaskGroups.Add "Overallocated Tasks", "Overallocated" 
ActiveProject.TaskGroups("Overallocated Tasks").GroupCriteria(1).Ascending = False
```


## Methods



|**Name**|
|:-----|
|[Add](taskgroups-add-method-project.md)|
|[Copy](taskgroups-copy-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](taskgroups-application-property-project.md)|
|[Count](taskgroups-count-property-project.md)|
|[Item](taskgroups-item-property-project.md)|
|[Parent](taskgroups-parent-property-project.md)|

