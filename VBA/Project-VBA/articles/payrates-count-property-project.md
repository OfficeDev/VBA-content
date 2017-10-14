---
title: PayRates.Count Property (Project)
ms.prod: project-server
api_name:
- Project.PayRates.Count
ms.assetid: 8ce675fe-ace2-b509-05dd-476c0d228bc1
ms.date: 06/08/2017
---


# PayRates.Count Property (Project)

Gets the number of items in the  **PayRates** collection. Read-only **Long**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **PayRates** object.


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


## See also


#### Concepts


[PayRates Collection Object](payrates-object-project.md)
