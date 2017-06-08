---
title: Month.Working Property (Project)
ms.prod: project-server
api_name:
- Project.Month.Working
ms.assetid: 6fa33218-2cf0-dbe4-af31-514c7c83a047
ms.date: 06/08/2017
---


# Month.Working Property (Project)

 **True** if any day in the month is a working day. Read/write **Boolean**.


## Syntax

 _expression_. **Working**

 _expression_ A variable that represents a **Month** object.


## Example

The following example makes June, July, and August nonworking months for resources in the "Student" group of the active project.


```vb
Sub GiveStudentsSummerOff() 
 
 Dim R As Resource ' Resource object used in For Each loop 
 Dim Y As Year ' Year object used in For Each loop 
 
 ' Look for resources in the "Student" group of the active project. 
 For Each R In ActiveProject.Resources 
 
 ' Give the summer off to resources in the "Student" group. 
 If R.Group = "Student" Then 
 For Each Y In R.Calendar.Years 
 Y.Months("June").Working = False 
 Y.Months("July").Working = False 
 Y.Months("August").Working = False 
 Next Y 
 End If 
 
 Next R 
 
End Sub
```


