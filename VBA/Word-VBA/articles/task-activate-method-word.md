---
title: Task.Activate Method (Word)
keywords: vbawd10.chm159514634
f1_keywords:
- vbawd10.chm159514634
ms.prod: word
api_name:
- Word.Task.Activate
ms.assetid: 99c9750a-35f4-ac84-649b-fc8788dc0904
ms.date: 06/08/2017
---


# Task.Activate Method (Word)

Activates the  **Task** object.


## Syntax

 _expression_ . **Activate**( **_Wait_** )

 _expression_ Required. A variable that represents a **[Task](task-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wait_|Optional| **Variant**| **True** to wait until the user has activated Word before activating the task. **False** to immediately activate the task, even if Word isn't active.|

## Example

This example activates the Notepad application if Notepad is in the  **Tasks** collection.


```vb
Sub ActivateNotePad() 
 Dim Task1 'Notepad must be open and in the Task List. 
 
 For Each Task1 In Tasks 
 If InStr(Task1.Name, "Notepad") > 0 Then 
 Task1.Activate 
 Task1.WindowState = wdWindowStateNormal 
 End If 
 Next Task1 
End Sub
```


## See also


#### Concepts


[Task Object](task-object-word.md)

