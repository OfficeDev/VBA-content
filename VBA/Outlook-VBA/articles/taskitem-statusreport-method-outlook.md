---
title: TaskItem.StatusReport Method (Outlook)
keywords: vbaol11.chm1757
f1_keywords:
- vbaol11.chm1757
ms.prod: outlook
api_name:
- Outlook.TaskItem.StatusReport
ms.assetid: 70549833-3287-bbbe-6756-896d400f6695
ms.date: 06/08/2017
---


# TaskItem.StatusReport Method (Outlook)

Sends a status report to all Cc recipients (recipients returned by the  **[StatusUpdateRecipients](taskitem-statusupdaterecipients-property-outlook.md)** property) with the current status for the task and returns an **Object** representing the status report.


## Syntax

 _expression_ . **StatusReport**

 _expression_ A variable that represents a **TaskItem** object.


### Return Value

An  **Object** value that represents the status report.


## Example

This Visual Basic for Applications (VBA) example uses the  **[StatusReport](taskitem-statusreport-method-outlook.md)** method to report the status of the currently open task.


```vb
Sub SendStatusReport() 
 Dim myTask As Outlook.TaskItem 
 Dim myinspector As Outlook.Inspector 
 Dim myReport As Object 
 
 Set myinspector = Application.ActiveInspector 
 If Not TypeName(myinspector) = "Nothing" Then 
 If TypeName(myinspector.CurrentItem) = "TaskItem" Then 
 Set myTask = myinspector.CurrentItem 
 Set myReport = myTask.StatusReport 
 myReport.Send 
 Else 
 MsgBox "No task item is currently open." 
 End If 
 Else 
 MsgBox "No inspector is currently open." 
 End If 
End Sub
```


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

