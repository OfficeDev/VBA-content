---
title: Task.HyperlinkAddress Property (Project)
ms.prod: project-server
api_name:
- Project.Task.HyperlinkAddress
ms.assetid: 0fd6c70e-df9e-1d6e-df65-aa1de2f98b44
ms.date: 06/08/2017
---


# Task.HyperlinkAddress Property (Project)

Gets or sets the URL or UNC path of a document. Read/write  **String**.


## Syntax

 _expression_. **HyperlinkAddress**

 _expression_ A variable that represents a **Task** object.


## Example

The following example adds a hyperlink to all tasks in the active project, including tasks in subprojects


```vb
Sub AddHyperlink() 
 Dim T As Task 
 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 T.Hyperlink = "Microsoft" 
 T.HyperlinkAddress = "http://www.microsoft.com/" 
 End If 
 Next T 
 
End Su
```


