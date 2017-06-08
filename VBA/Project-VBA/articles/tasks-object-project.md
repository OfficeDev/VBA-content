---
title: Tasks Object (Project)
ms.prod: project-server
ms.assetid: b7482b5a-7fac-531e-6793-610faca2f954
ms.date: 06/08/2017
---


# Tasks Object (Project)

Contains a collection of  **[Task](task-object-project.md)** objects.


## Example

 **Using the Task Object**

Use  **Tasks** ( _Index_ ), where _Index_ is the task index number or task name, to return a single **Task** object. The following example prints the names of every resource assigned to every task in the active project.




```
Dim Temp As Long, A As Assignment 

Dim TaskName As String, Assigned As String, Results As String 

 

For Temp = 1 To ActiveProject.Tasks.Count 

 TaskName = "Task: " &amp; ActiveProject.Tasks(Temp).Name &amp; vbCrLf 

 For Each A In ActiveProject.Tasks(Temp).Assignments 

 Assigned = A.ResourceName &amp; ListSeparator &amp; " " &amp; Assigned 

 Next A 

 Results = Results &amp; TaskName &amp; "Resources: " &amp; _ 

 Left$(Assigned, Len(Assigned) - Len(ListSeparator &amp; " ")) &amp; vbCrLf &amp; vbCrLf 

 TaskName = "" 

 Assigned = "" 

Next Temp 

 

MsgBox Results
```

Use the  **[Tasks](http://msdn.microsoft.com/library/8f58ea8e-a3a1-f5aa-ad5d-6447fe777453%28Office.15%29.aspx)** property to return a **Tasks** collection. The following example displays the name of every task in the selection.




```
Dim T As Task, Names As String 

 

For Each T In ActiveSelection.Tasks 

 Names = Names &amp; T.Name &amp; vbCrLf 

Next T 

 

MsgBox Names
```

Use the  **[Add](http://msdn.microsoft.com/library/a6e2186b-610c-0888-a22a-8b7deba3f53f%28Office.15%29.aspx)** method to add a **Task** object to the **Tasks** collection. The following example adds a new task to the end of the task list.




```
ActiveProject.Tasks.Add "Hang clocks"
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/a6e2186b-610c-0888-a22a-8b7deba3f53f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/0d4405af-9edd-f8ad-b0ac-d72e0f02b16c%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/23238c44-1cf0-8dfc-40b3-6def228d5a7a%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/2bbdddae-38f7-6740-0694-73e0cf838e90%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/a2e8cfce-9c04-6c1f-badc-0fe506df270b%28Office.15%29.aspx)|
|[UniqueID](http://msdn.microsoft.com/library/f87b88e3-5bd0-a57b-c54b-aba17d0de67e%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
