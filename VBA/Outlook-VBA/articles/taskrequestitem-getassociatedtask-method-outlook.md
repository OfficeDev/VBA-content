---
title: TaskRequestItem.GetAssociatedTask Method (Outlook)
keywords: vbaol11.chm1906
f1_keywords:
- vbaol11.chm1906
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.GetAssociatedTask
ms.assetid: ec170266-9898-79d8-03e9-7ea38d789d40
ms.date: 06/08/2017
---


# TaskRequestItem.GetAssociatedTask Method (Outlook)

Returns a  **[TaskItem](taskitem-object-outlook.md)** object that represents the requested task.


## Syntax

 _expression_ . **GetAssociatedTask**( **_AddToTaskList_** )

 _expression_ A variable that represents a **TaskRequestItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AddToTaskList_|Required| **Boolean**| **True** if the task is added to the default **Tasks** folder.|

### Return Value

A  **TaskItem** object that represents the requested task.


## Remarks

The  **GetAssociatedTask** method will not work unless the **TaskItem** is processed before the method is called. To do so, call the **[Display](taskitem-display-method-outlook.md)** method before calling **GetAssociatedTask** .


## Example

This Microsoft Visual Basic for Applications (VBA) example accepts a  **[TaskRequestItem](taskrequestitem-object-outlook.md)** , sending the response without displaying the inspector.


```vb
Sub AcceptTask() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myTasks As Outlook.Folder 
 
 Dim myNewTaskItem As Outlook.TaskItem 
 
 Dim mytaskreqItem As Outlook.TaskRequestItem 
 
 Dim myItem As Outlook.TaskItem 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myTasks = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set mytaskreqItem = myTasks.Items.Find("[Subject] = ""Meeting w/ Nate Sun""") 
 
 If Not TypeName(mytaskreqItem) = "Nothing" Then 
 
 Set myNewTaskItem = mytaskreqItem.GetAssociatedTask(True) 
 
 Set myItem = myNewTaskItem.Respond(olTaskAccept, True, True) 
 
 myItem.Send 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[TaskRequestItem Object](taskrequestitem-object-outlook.md)

