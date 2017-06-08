---
title: StartDriver.Suggestions Property (Project)
ms.prod: project-server
api_name:
- Project.StartDriver.Suggestions
ms.assetid: 39cfa3ae-ca39-7260-ebe4-a0abe40b3799
ms.date: 06/08/2017
---


# StartDriver.Suggestions Property (Project)

Gets a combination of  **[PjTaskWarnings](pjtaskwarnings-enumeration-project.md)** values that indicate whether there are potential problems that should be fixed for a specified task. Read-only **Long**.


## Syntax

 _expression_. **Suggestions**

 _expression_ An expression that returns a **StartDriver** object.


## Remarks

If there are no suggestions for a task, the value of  **Suggestions** is 0. Because the value of **pjTaskWarningsResourceBeyondMaxUnit** is 64 and the value of **pjTaskWarningsResourceOverallocated** is 128, if **Suggestions** is 192, the task has both of the potential problems.


 **Note**  The  **PjTaskWarnings** enumeration can be used with both the **Suggestions** property and the **[Warnings](startdriver-warnings-property-project.md)** property.


## Example

In the following example, if the value of the  **Suggestions** property for task 2 is 128, the message box shows **The resource is overallocated**. If the value is 68, the message box shows:


-  **The assignment is more than the maximum resource units available.**
    
-  **The shadow task finishes earlier because of a predecessor link.**
    





```vb
Sub GetTaskSuggestions() 

 Dim suggestions As Long 

 Dim suggestionMsg As String 

 

 suggestions = ActiveProject.Tasks(2).StartDriver.Suggestions 

 

 suggestionMsg = CheckSuggestions(suggestions) 

 

 If Not suggestionMsg = "" Then MsgBox suggestionMsg 

End Sub 

 

Function CheckSuggestions(suggestions As Long) As String 

 Dim partial As Long 

 Dim suggestionResult As String 

 

 suggestionResult = "" 

 partial = suggestions Xor pjTaskWarningResourceBeyondMaxUnit 

 If partial < suggestions Then _ 

 suggestionResult = suggestionResult &; "The assignment is more than the maximum resource units available." &; vbCrLf 

 

 partial = suggestions Xor pjTaskWarningResourceOverallocated 

 If partial < suggestions Then _ 

 suggestionResult = suggestionResult &; "The resource is overallocated." &; vbCrLf 

 

 partial = suggestions Xor pjTaskWarningShadowFinishesEarlierDueToLink 

 If partial < suggestions Then _ 

 suggestionResult = suggestionResult &; "The shadow task finishes earlier because of a predecessor link." &; vbCrLf 

 

 CheckSuggestions = suggestionResult 

End Function
```


## See also


#### Concepts


[StartDriver Object](startdriver-object-project.md)
