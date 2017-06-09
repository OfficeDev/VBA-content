---
title: TaskItem.Send Method (Outlook)
keywords: vbaol11.chm1755
f1_keywords:
- vbaol11.chm1755
ms.prod: outlook
api_name:
- Outlook.TaskItem.Send
ms.assetid: 54f751fc-cff1-5d17-f635-f688cd8ad6f8
ms.date: 06/08/2017
---


# TaskItem.Send Method (Outlook)

Sends the task.


## Syntax

 _expression_ . **Send**

 _expression_ A variable that represents a **TaskItem** object.


## Example

This Visual Basic for Applications (VBA) example uses  **[CreateItem](application-createitem-method-outlook.md)** to create a simple task and delegate it as a task request to another user. Replace 'Dan Wilson' with a valid recipient name before running this example.


```vb
Sub AssignTask() 
 
 Dim myItem As Outlook.TaskItem 
 
 Dim myDelegate As Outlook.Recipient 
 
 
 
 Set MyItem = Application.CreateItem(olTaskItem) 
 
 MyItem.Assign 
 
 Set myDelegate = MyItem.Recipients.Add("Dan Wilson") 
 
 myDelegate.Resolve 
 
 If myDelegate.Resolved Then 
 
 myItem.Subject = "Prepare Agenda for Meeting" 
 
 myItem.DueDate = Now + 30 
 
 myItem.Display 
 
 myItem.Send 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

