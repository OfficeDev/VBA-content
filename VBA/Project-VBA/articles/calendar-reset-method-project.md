---
title: Calendar.Reset Method (Project)
keywords: vbapj.chm131258
f1_keywords:
- vbapj.chm131258
ms.prod: project-server
api_name:
- Project.Calendar.Reset
ms.assetid: fc638f47-36b5-aa36-55c2-882bd570b9cb
ms.date: 06/08/2017
---


# Calendar.Reset Method (Project)

Resets base calendar properties to their default values; resets resource calendar properties to the values in the corresponding base calendar.


## Syntax

 _expression_. **Reset**

 _expression_ A variable that represents a **Calendar** object.


## Example

The following example resets every resource calendar in the active project.


```vb
Sub ResetResourceCalendars() 
 
 Dim R As Resource ' Resource object used in For Each loop 
 
 For Each R In ActiveProject.Resources 
 R.Calendar.Reset 
 Next R 
 
End Sub
```


