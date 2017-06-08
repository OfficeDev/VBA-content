---
title: OverAllocatedAssignments.Item Property (Project)
ms.prod: project-server
api_name:
- Project.OverAllocatedAssignments.Item
ms.assetid: 5939e712-0abd-cb4b-31fe-ad2fa61835d6
ms.date: 06/08/2017
---


# OverAllocatedAssignments.Item Property (Project)

Gets a single  **Assignment** object from the **OverAllocatedAssignments** collection. Read-only **[Assignment](assignment-object-project.md)**.


## Syntax

 _expression_. **Item**( ** _Index_** )

 _expression_ An expression that returns an **OverAllocatedAssignments** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The index number of the  **Assignment** to return.|

## Example

The following example finds assignments where the resource is overallocated. When the overPeak argument is  **False**, the overallocation is not greater than the maximum resource time available (100%). If you set overPeak to **True**, the example finds overallocated assignments that exceed maximum resource time available, such as 150%.


```vb
Sub FindOverallocatedAssignments() 

 Dim t As Task 

 Dim a As Assignment 

 Dim overAlloc As OverAllocatedAssignments 

 Dim numOver As Long 

 Dim i As Long 

 Dim overPeak As Boolean 

 

 overPeak = True 

 

 For Each t In ActiveProject.Tasks 

 If t.Overallocated Then 

 Set overAlloc = t.StartDriver.OverAllocatedAssignments(overPeak) 

 numOver = overAlloc.Count 

 totalNumOver = overAlloc.TotalDetectedCount 

 

 For i = 1 To numOver 

 Set a = overAlloc.Item(i) 

 Debug.Print "Task: " &; t.Name &; " - Overallocated resource: " _ 

 &; a.ResourceName 

 Debug.Print vbTab &; "Resource peak: " &; a.Peak 

 Next i 

 End If 

 Next t 

End Sub
```


## See also


#### Concepts


[OverAllocatedAssignments Collection Object](overallocatedassignments-object-project.md)

