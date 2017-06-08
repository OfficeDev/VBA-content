---
title: Resource.Group Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Group
ms.assetid: 9f5f5bd6-c104-629c-feab-455fbeaf27eb
ms.date: 06/08/2017
---


# Resource.Group Property (Project)

Gets or sets the group to which a resource belongs. Read/write  **String**.


## Syntax

 _expression_. **Group**

 _expression_ A variable that represents a **Resource** object.


## Example

The following example deletes the resources in the active project that belong to a group specified by the user.


```vb
Sub DeleteResourcesInGroup() 
 
 Dim Entry As String ' The group specified by the user 
 Dim Deletions As Integer ' The number of deleted resources 
 Dim R As Resource ' The resource object used in loop 
 
 ' Prompt user for the name of a group. 
 Entry = InputBox$("Enter a group name:") 
 
 ' Cycle through the resources of the active project. 
 For Each R in ActiveProject.Resources 
 ' Delete a resource if its group name matches the user's request. 
 If R.Group = Entry Then 
 R.Delete 
 Deletions = Deletions + 1 
 End If 
 Next R 
 
 ' Display the number of resources that were deleted. 
 MsgBox(Deletions &; " resources were deleted.") 
 
End Sub
```


