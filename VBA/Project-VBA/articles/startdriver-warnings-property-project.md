---
title: StartDriver.Warnings Property (Project)
ms.prod: project-server
api_name:
- Project.StartDriver.Warnings
ms.assetid: 409c84e2-8307-fb82-af19-fa0e9f6b406b
ms.date: 06/08/2017
---


# StartDriver.Warnings Property (Project)

Gets a combination of  **[PjTaskWarnings](pjtaskwarnings-enumeration-project.md)** values that indicate whether there are problems for a specified task. Read-only **Long**.


## Syntax

 _expression_. **Warnings**

 _expression_ An expression that returns a **StartDriver** object.


## Remarks

If there are no warnings for a task, the value of  **Warnings** is 0. Because the value of **pjTaskWarningResourceBeyondMaxUnit** is 64 and the value of **pjTaskWarningResourceOverallocated** is 128, if **Warnings** is 192, the task has both of the problems.


 **Note**  The  **PjTaskWarnings** enumeration can be used with both the **[Suggestions](startdriver-suggestions-property-project.md)** property and the **Warnings** property.


## Example

In the following example, if the value of the  **Warnings** property for task 5 is 128, the message box shows **The resource is overallocated.**. If the value is 68, the message box shows:


-  **The assignment is more than the maximum resource units available.**
    
-  **The shadow task finishes earlier because of a predecessor link.**
    





```vb
Sub GetTaskWarnings() 

 Dim warnings As Long 

 Dim warningMsg As String 

 

 warnings = ActiveProject.Tasks(5).StartDriver.Warnings 

 

 warningMsg = CheckWarnings(warnings) 

 

 If Not warningMsg = "" Then MsgBox warningMsg 

End Sub 

 

Function CheckWarnings(warnings As Long) As String 

 Dim partial As Long 

 Dim warningResult As String 

 

 warningResult = "" 

 partial = warnings Xor pjTaskWarningResourceBeyondMaxUnit 

 If partial < warnings Then _ 

 warningResult = warningResult &; "The assignment is more than the maximum resource units available." &; vbCrLf 

 

 partial = warnings Xor pjTaskWarningResourceOverallocated 

 If partial < warnings Then _ 

 warningResult = warningResult &; "The resource is overallocated." &; vbCrLf 

 

 partial = warnings Xor pjTaskWarningShadowFinishesEarlierDueToLink 

 If partial < warnings Then _ 

 warningResult = warningResult &; "The shadow task finishes earlier because of a predecessor link." &; vbCrLf 

 

 CheckWarnings = warningResult 

End Function
```


## See also


#### Concepts


[StartDriver Object](startdriver-object-project.md)
