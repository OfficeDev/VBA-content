---
title: List.Count Property (Project)
ms.prod: project-server
api_name:
- Project.List.Count
ms.assetid: d417857d-99f9-3c82-f211-4dd0241deb44
ms.date: 06/08/2017
---


# List.Count Property (Project)

Gets the number of items in the  **List** collection. Read-only **Long**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **List** object.


## Example

The following example prompts the user for the name of a resource and then assigns that resource to tasks without any resources.


```vb
Sub AssignResource() 
 
 Dim T As Task ' Task object used in For Each loop 
 Dim R As Resource ' Resource object used in For Each loop 
 Dim Rname As String ' Resource name 
 Dim RID As Long ' Resource ID 
 
 RID = 0 
 RName = InputBox$("Enter the name of a resource: ") 
 
 For Each R in ActiveProject.Resources 
 If R.Name = RName Then 
 RID = R.ID 
 Exit For 
 End If 
 Next R 
 
 If RID <> 0 Then 
 ' Assign the resource to tasks without any resources. 
 For Each T In ActiveProject.Tasks 
 If T.Assignments.Count = 0 Then 
 T.Assignments.Add ResourceID:=RID 
 End If 
 Next T 
 Else 
 MsgBox Prompt:=RName &; " is not a resource in this project.", buttons:=vbExclamation 
 End If 
 
End Sub
```


