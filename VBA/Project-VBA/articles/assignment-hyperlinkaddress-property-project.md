---
title: Assignment.HyperlinkAddress Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.HyperlinkAddress
ms.assetid: ead317d6-aa1a-57a1-4d58-189ccf551b40
ms.date: 06/08/2017
---


# Assignment.HyperlinkAddress Property (Project)

Gets or sets the URL or UNC path of a document. Read/write  **String**.


## Syntax

 _expression_. **HyperlinkAddress**

 _expression_ A variable that represents an **Assignment** object.


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


