---
title: SplitPart Object (Project)
ms.prod: project-server
api_name:
- Project.SplitPart
ms.assetid: 7eb80010-7b5a-3833-a5c5-b02d0c0bea5c
ms.date: 06/08/2017
---


# SplitPart Object (Project)

Represents a task portion. The  **SplitPart** object is a member of the **[SplitParts](splitparts-object-project.md)** collection.
 


## Examples

 **Using the SplitPart Object**
 

 
Use  **SplitParts** (*Index* ), where*Index* is the index number of the task portion, to return a single **SplitPart** object. The following example lists the start and finish times of each task portion of the task in the active cell.
 

 



```
Dim Part As Long, Portions As String

For Part = 1 To ActiveCell.Task.SplitParts.Count
    With ActiveCell.Task
        Portions = Portions &amp; "Task portion " &amp; Part &amp; ": Start on " &amp; _
            .SplitParts(Part).Start &amp; ", Finish on " &amp; _
            .SplitParts(Part).Finish &amp; vbCrLf
    End With
Next Part

MsgBox Portions
```

 **Using the SplitParts Collection**
 

 
Use the  **[SplitParts](task-splitparts-property-project.md)** property to return a **SplitParts** collection. The following example returns the number of task portions for each task in the active project.
 

 



```
Dim T As Task

For Each T In ActiveProject.Tasks
    If Not (T Is Nothing) Then
        MsgBox T.Name &amp; ": " &amp; T.SplitParts.Count
    End If

Next T
```

Use the  **[Split](task-split-method-project.md)** method ( **Task** object) to add a **SplitPart** object to the **SplitParts** collection. (The **Split** method creates a split in a task.) The following example creates a split in the task from Wednesday to Monday, in October of 2012.
 

 



```
ActiveCell.Task.Split "10/3/2012", "10/8/2012"
```


## Methods



|**Name**|
|:-----|
|[Delete](splitpart-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](splitpart-application-property-project.md)|
|[Finish](splitpart-finish-property-project.md)|
|[Index](splitpart-index-property-project.md)|
|[Parent](splitpart-parent-property-project.md)|
|[Start](splitpart-start-property-project.md)|

