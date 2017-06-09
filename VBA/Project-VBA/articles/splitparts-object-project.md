---
title: SplitParts Object (Project)
ms.prod: project-server
ms.assetid: bc36310c-9289-a363-f2d6-c8a0991725e5
ms.date: 06/08/2017
---


# SplitParts Object (Project)

Contains a collection of  **[SplitPart](splitpart-object-project.md)** objects.
 


## Example

 **Using the SplitParts Collection Object**
 

 
Use  **SplitParts** (*Index* ), where*Index* is the index number of the task index number, to return a single **SplitPart** object. The following example lists the start and finish times of each task portion of the task in the active cell.
 

 



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

Use the  **[Split](task-split-method-project.md)** method ( **Task** object) to add a **SplitPart** object to the **SplitParts** collection. (The **Split** method creates a split in a task.) The following example creates a split in the task from Wednesday to Monday.
 

 



```
ActiveCell.Task.Split "10/2/02", "10/7/02"
```


## Methods



|**Name**|
|:-----|
|[Add](splitparts-add-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](splitparts-application-property-project.md)|
|[Count](splitparts-count-property-project.md)|
|[Item](splitparts-item-property-project.md)|
|[Parent](splitparts-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
