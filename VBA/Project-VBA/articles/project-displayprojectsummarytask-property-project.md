---
title: Project.DisplayProjectSummaryTask Property (Project)
keywords: vbapj.chm131748
f1_keywords:
- vbapj.chm131748
ms.prod: PROJECTSERVER
api_name:
- Project.Project.DisplayProjectSummaryTask
ms.assetid: 4b04ec4a-a050-8038-c549-bc8942fbadd6
---


# Project.DisplayProjectSummaryTask Property (Project)

 **True** if the summary task for a project is visible. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayProjectSummaryTask**

 _expression_ A variable that represents a **Project** object.


## Example

The following example creates a new project and displays its summary task.


```vb
Sub NewProject() 
 
 FileNew 
 ActiveProject.DisplayProjectSummaryTask = True 
 
End Sub
```


