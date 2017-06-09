---
title: TaskGroups2 Object (Project)
ms.prod: project-server
ms.assetid: 08346fd5-3dbd-23ea-9dc8-c2361ce043f4
ms.date: 06/08/2017
---


# TaskGroups2 Object (Project)

Represents all the task-based group definitions, where group hierarchy can be maintained.  **TaskGroups2** is a collection of **[Group2](group2-object-project.md)** objects.
 


## Example

 **Using the TaskGroups2 Collection**
 

 
Use the  **[TaskGroups2](project-taskgroups2-property-project.md)** property to return a **TaskGroups2** collection. The following example lists the names of all the task groups in the active project.
 

 



```
Dim tg2 As Group2
Dim tGroups2 As String

For Each tg2 in ActiveProject.TaskGroups2  
    tGroups2 = tGroups2 &amp; tg2.Name &amp; vbCrLf  
Next tg2  

MsgBox tGroups2
```

Use the  **[Add](taskgroups2-add-method-project.md)** method to add a **Group2** object to the **TaskGroups2** collection. The following example creates a new group that groups tasks by whether they are overallocated and then modifies the criterion so that overallocated tasks are sorted in descending order.
 

 



```
ActiveProject.TaskGroups2.Add "Overallocated Tasks", "Overallocated"
ActiveProject.TaskGroups2("Overallocated Tasks").GroupCriteria(1).Ascending = False
```


## Methods



|**Name**|
|:-----|
|[Add](taskgroups2-add-method-project.md)|
|[Copy](taskgroups2-copy-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](taskgroups2-application-property-project.md)|
|[Count](taskgroups2-count-property-project.md)|
|[Item](taskgroups2-item-property-project.md)|
|[Parent](taskgroups2-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
