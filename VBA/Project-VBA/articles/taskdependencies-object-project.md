---
title: TaskDependencies Object (Project)
ms.prod: project-server
ms.assetid: 60bda111-998f-1cc2-0b18-b419041767f5
ms.date: 06/08/2017
---


# TaskDependencies Object (Project)

Contains a collection of  **[TaskDependency](taskdependency-object-project.md)** objects.


## Example

 **Using the TaskDependency Object**

Use  **TaskDependencies** ( _Index_ ), where _Index_ is the dependency index, to return a single **TaskDependency** object. The following example adds 1.5 days of lag to the link between the specified task and the predecessor specified in its first task dependency.




```
ActiveProject.Tasks("Draft Initial Business Case").TaskDependencies(1).Lag = "1.5d"
```

 **Using the TaskDependencies Collection**

Use the  **[TaskDependencies](http://msdn.microsoft.com/library/9c02fe5f-cb9e-a10e-bf9a-66b7600f8c64%28Office.15%29.aspx)** property to return a **TaskDependencies** collection. The following example examines each predecessor for the specified task and displays a message for each that has a priority of "High" or better.




```
Dim TaskDep As TaskDependency 

 

For Each TaskDep In ActiveProject.Tasks("Write Requirements Brief").TaskDependencies 

 If TaskDep.From.Priority > 500 Then 

 MsgBox "Task #" &amp; TaskDep.From.ID &amp; " (" &amp; TaskDep.From.Name &amp; ") " &amp; _ 

 "has a priority higher than medium." 

 End If 

Next TaskDep
```

Use the  **[Add](http://msdn.microsoft.com/library/37e67ab2-ca7b-26c2-50e7-8a933b746489%28Office.15%29.aspx)** method to add a **TaskDependency** object to the **TaskDependencies** collection. The following example links "Preliminary Research &amp; Approval" as a predecessor to "Draft Initial Business Case" in a finish-to-start relationship.




```
ActiveProject.Tasks("Draft Initial Business Case").TaskDependencies.Add ActiveProject.Tasks("Preliminary Research &amp; Approval"), pjFinishToStart
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/37e67ab2-ca7b-26c2-50e7-8a933b746489%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/8eccf4cd-a1d6-8e8b-b5e4-c5a3f43463eb%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/499ae3c9-b99a-be2b-2d57-7f3dcb28d683%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/b43d6c70-ee9a-d022-93cf-696725d48fd8%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/965a8751-4e56-9846-a1b6-d83163f5dfef%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
