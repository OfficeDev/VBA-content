---
title: Project.LevelToDate Property (Project)
ms.prod: project-server
api_name:
- Project.Project.LevelToDate
ms.assetid: b697db71-8f8e-9caf-345c-59899f4024a3
ms.date: 06/08/2017
---


# Project.LevelToDate Property (Project)

Gets or sets the ending date of a range in which overallocated resources are leveled. The default is the project finish date or the last entered date value. Read/write  **Variant**.


## Syntax

 _expression_. **LevelToDate**

 _expression_ A variable that represents a **Project** object.


## Remarks

You can also set the  **LevelToDate** property in the **Resource Leveling** dialog box. To access the setting, click **Leveling Options** on the **Resource** tab of the Ribbon, and then click the **Level** option and set the **To** date.


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


