---
title: OverAllocatedAssignments Object (Project)
ms.prod: project-server
ms.assetid: b2856ebf-cff2-04a6-53c9-123de09f2a3b
ms.date: 06/08/2017
---


# OverAllocatedAssignments Object (Project)

Represents a collection of  **[Assignment](assignment-object-project.md)** objects where the resource is overallocated.
 


## Remarks

Use the  **[Item](overallocatedassignments-item-property-project.md)** property to get a single **Assignment** object from the **OverAllocatedAssignments** collection.
 

 

## Example

The following example finds assignments where the resource is overallocated. When the overPeak argument is  **False**, the overallocation is not greater than the maximum resource time available (100%). If you set overPeak to **True**, the example finds overallocated assignments that exceed maximum resource time available, such as 150%.
 

 

```
Sub FindOverallocatedAssignments()  
    Dim t As Task  
    Dim a As Assignment  
    Dim overAlloc As OverAllocatedAssignments  
    Dim numOver As Long  
    Dim overPeak As Boolean  
  
    overPeak = False  
  
    For Each t In ActiveProject.Tasks  
        If t.Overallocated Then  
            Set overAlloc = t.StartDriver.OverAllocatedAssignments(overPeak)  
            numOver = overAlloc.Count  
            totalNumOver = overAlloc.TotalDetectedCount  
  
            For Each a In overAlloc  
                Debug.Print "Resource: " &amp; a.Resource.Name &amp; " is overallocated on task: " &amp; t.Name  
                Debug.Print vbTab &amp; "Number of overallocated assignments: " &amp; numOver  
            Next a  
        End If  
    Next t  
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](overallocatedassignments-application-property-project.md)|
|[Count](overallocatedassignments-count-property-project.md)|
|[Item](overallocatedassignments-item-property-project.md)|
|[Parent](overallocatedassignments-parent-property-project.md)|
|[TotalDetectedCount](overallocatedassignments-totaldetectedcount-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
