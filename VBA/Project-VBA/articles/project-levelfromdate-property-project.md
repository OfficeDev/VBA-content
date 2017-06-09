---
title: Project.LevelFromDate Property (Project)
ms.prod: project-server
api_name:
- Project.Project.LevelFromDate
ms.assetid: 19e29259-de9d-9e8a-b724-129839dca23b
ms.date: 06/08/2017
---


# Project.LevelFromDate Property (Project)

Gets or sets the starting date of a range in which overallocated resources are leveled. The default is the project start date or the last entered date value. Read/write  **Variant**.


## Syntax

 _expression_. **LevelFromDate**

 _expression_ A variable that represents a **Project** object.


## Remarks

You can also set the  **LevelFromDate** property in the **Resource Leveling** dialog box. To access the setting, click **Leveling Options** on the **Resource** tab of the Ribbon, and then click the **Level** option and set the **From** date.


## Example

The following example lets the user change the date range where leveling occurs, if the current range starts on the project start date or finishes on the project finish date.


```vb
Sub ChangeLevelingDates() 
 Dim Response As Long 
 Dim NewFrom As Variant, NewTo As Variant 
 
 With ActiveProject 
 If Application.DateDifference(.ProjectSummaryTask.Start, .LevelFromDate) = 0 Then 
 Response = MsgBox("Overallocated resources are leveled from " &; _ 
 "the project start date. Should that be changed?", vbYesNo) 
 If Response = vbYes Then 
 NewFrom = InputBox("Date to level from: ") 
 .LevelFromDate = NewFrom 
 Else 
 MsgBox "Resources remain leveled from the project start date." 
 End If 
 End If 
 
 If Application.DateDifference(.ProjectSummaryTask.Finish, .LevelToDate) = 0 Then 
 Response = MsgBox("Overallocated resources are leveled to " &; _ 
 "the project finish date. Should that be changed?", vbYesNo) 
 If Response = vbYes Then 
 NewTo = InputBox("Date to level to: ") 
 .LevelToDate = NewTo 
 Else 
 MsgBox "Resources remain leveled to the project finish date." 
 End If 
 End If 
 End With 
 
End Sub
```


