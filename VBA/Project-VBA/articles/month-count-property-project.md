---
title: Month.Count Property (Project)
ms.prod: project-server
api_name:
- Project.Month.Count
ms.assetid: cf17523e-9b43-ee38-3c45-15936e8d0559
ms.date: 06/08/2017
---


# Month.Count Property (Project)

Gets the number of days in the  **Month** object. Read-only **Integer**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Month** object.


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


