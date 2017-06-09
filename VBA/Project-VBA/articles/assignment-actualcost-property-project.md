---
title: Assignment.ActualCost Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.ActualCost
ms.assetid: 45bf4d44-bce7-474a-7093-ff0c97d3b7f6
ms.date: 06/08/2017
---


# Assignment.ActualCost Property (Project)

Gets or sets the actual cost for the assignment. Read/write  **Variant**.


## Syntax

 _expression_. **ActualCost**

 _expression_ A variable that represents an **Assignment** object.


## Remarks

The  **ActualCost** property can be set for **Assignment** and **Task** objects (but not for summary tasks) if the **Actual costs are always calculated by Project** check box is cleared on the **Schedule** tab of the **Project Options** dialog box.

Actual costs are also available for tasks and resources. If the  **Actual costs are always calculated by Project** check box is checked, Project calculates the current actual cost for the assignment from resource cost rate tables and actual work that the assigned resource has completed. For programmatic access to the resource cost rate tables, use the **[CostRateTables](resource-costratetables-property-project.md)** collection.


## Example

The following example prompts the user for actual costs of tasks with no resources in the active project. It assumes that the  **Actual costs are always calculated by Project** check box is cleared.


```vb
Sub GetActualCostsForTasks() 
 
 Dim Entry As String ' User input 
 Dim T As Task ' Task object used in For Each loop 
 
 ' Count the resources of each task in the active project. 
 For Each T In ActiveProject.Tasks 
 
 ' If a task has no resources, then prompt user for actual cost. 
 If T.Resources.Count = 0 Then 
 
 Do While 1 
 Entry = InputBox$("Enter the cost for " &; T.Name &; ":") 
 
 ' Exit loop if user enters number or clicks Cancel. 
 If IsNumeric(Entry) Or Entry = Empty Then 
 Exit Do 
 
 ' User didn't enter a number; tell user to try again. 
 Else 
 MsgBox ("You didn't enter a number; try again.") 
 End If 
 Loop 
 
 ' If user didn't click Cancel, assign actual cost to task. 
 If Not StrComp(Entry, Empty, 1) = 0 Then T.ActualCost = Entry 
 End If 
 
 Next T 
 
End Sub
```


