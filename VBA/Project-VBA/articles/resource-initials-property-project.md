---
title: Resource.Initials Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Initials
ms.assetid: b74494c1-955d-2984-9c3c-4271d382deb1
ms.date: 06/08/2017
---


# Resource.Initials Property (Project)

Gets or sets the initials of a resource. Read/write  **String**.


## Syntax

 _expression_. **Initials**

 _expression_ A variable that represents a **Resource** object.


## Example

The following example sets the initials of each resource in the active project according to the spaces in the resource's name. For example, a resource called "Glue Gun" receives the initials "GG."


```vb
Sub SetInitialsBasedOnName() 
 
 Dim I As Integer ' Index used in For loop 
 Dim R As Resource ' Resource used in For Each loop 
 Dim NewInits As String ' The new initials of the resource 
 
 ' Cycle through the resources of the active project. 
 For Each R In ActiveProject.Resources 
 ' Initialize with first character of name. 
 NewInits = Mid(R.Name, 1, 1) 
 
 ' Look for spaces in the resource's name. 
 For I = 1 To Len(R.Name) 
 ' If not first character, and space is found, then add initial. 
 If I > 1 And Mid(R.Name, I, 1) = Chr(32) Then 
 If I + 1 <= Len(R.Name) Then 
 NewInits = NewInits &; Mid(R.Name, I + 1, 1) 
 End If 
 End If 
 Next I 
 
 ' Give the resource its new initials. 
 R.Initials = NewInits 
 
 Next R 
 
End Sub
```


