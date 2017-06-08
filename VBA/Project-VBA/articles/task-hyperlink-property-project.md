---
title: Task.Hyperlink Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Hyperlink
ms.assetid: 72138942-5e4b-3fe5-abd1-f907b2ee315d
ms.date: 06/08/2017
---


# Task.Hyperlink Property (Project)

Gets or sets a friendly name representing a hyperlink address. The name may also be a URL or UNC path. Read/write  **String**.


## Syntax

 _expression_. **Hyperlink**

 _expression_ A variable that represents a **Task** object.


## Example

The following example adds a hyperlink to all tasks in the active project, including tasks in subprojects.


```vb
Sub AddHyperlink() 
 Dim T As Task 
 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 T.Hyperlink = "Microsoft" 
 T.HyperlinkAddress = "http://www.microsoft.com/" 
 End If 
 Next T 
 
End Sub
```


