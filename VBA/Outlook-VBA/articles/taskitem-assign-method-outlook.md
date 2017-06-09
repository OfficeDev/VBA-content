---
title: TaskItem.Assign Method (Outlook)
keywords: vbaol11.chm1749
f1_keywords:
- vbaol11.chm1749
ms.prod: outlook
api_name:
- Outlook.TaskItem.Assign
ms.assetid: f254107a-4182-de3a-2039-08f664e61eeb
ms.date: 06/08/2017
---


# TaskItem.Assign Method (Outlook)

Assigns a task and returns a  **[TaskItem](taskitem-object-outlook.md)** object that represents it.


## Syntax

 _expression_ . **Assign**

 _expression_ A variable that represents a **TaskItem** object.


### Return Value

A  **TaskItem** object that represents the task.


## Remarks

This method allows a task to be assigned (delegated) to another user. You must create a task before you can assign it, and you must assign a task before you can send it. An assigned task is sent as a  **[TaskRequestItem](taskrequestitem-object-outlook.md)** object.


## Example

This Visual Basic for Applications (VBA) example uses  **[CreateItem](application-createitem-method-outlook.md)** to create a simple task and delegate it as a task request to another user. To run this example, replace 'Dan Wilson' with a valid recipient name.


```vb
Sub AssignTask() 
 
 Dim myItem As Outlook.TaskItem 
 
 Dim myDelegate As Outlook.Recipient 
 
 
 
 Set MyItem = Application.CreateItem(olTaskItem) 
 
 MyItem.Assign 
 
 Set myDelegate = MyItem.Recipients.Add("Dan Wilson") 
 
 myDelegate.Resolve 
 
 If myDelegate.Resolved Then 
 
 myItem.Subject = "Prepare Agenda For Meeting" 
 
 myItem.DueDate = Now + 30 
 
 myItem.Display 
 
 myItem.Send 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

