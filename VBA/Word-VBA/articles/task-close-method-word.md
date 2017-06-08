---
title: Task.Close Method (Word)
keywords: vbawd10.chm159514635
f1_keywords:
- vbawd10.chm159514635
ms.prod: word
api_name:
- Word.Task.Close
ms.assetid: 455a51bd-90f5-b14b-497e-618fe4df5406
ms.date: 06/08/2017
---


# Task.Close Method (Word)

Closes the specified task.


## Syntax

 _expression_ . **Close**

 _expression_ Required. A variable that represents a **[Task](task-object-word.md)** object.


## Example

This example activates Microsoft Excel and then closes it.


```vb
For Each myTask In Tasks 
 If InStr(myTask.Name, "Microsoft Excel") > 0 Then 
 myTask.Activate 
 myTask.Close 
 End If 
Next myTask
```


## See also


#### Concepts


[Task Object](task-object-word.md)

