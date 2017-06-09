---
title: Tasks.Exists Method (Word)
keywords: vbawd10.chm159580162
f1_keywords:
- vbawd10.chm159580162
ms.prod: word
api_name:
- Word.Tasks.Exists
ms.assetid: 421a5ff6-25b5-3255-ae81-32f5decbfe93
ms.date: 06/08/2017
---


# Tasks.Exists Method (Word)

Determines whether the specified task exists. Returns  **True** if the task exists.


## Syntax

 _expression_ . **Exists**( **_Name_** )

 _expression_ A variable that represents a **[Tasks](tasks-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the task.|

### Return Value

Boolean


## Example

This example determines whether the Windows Calculator program is running (if the task exists). If Calculator isn't running, the Shell statement starts it. If Calculator is running, the application is activated.


```vb
If Tasks.Exists("Calculator") = False Then 
 Shell "Calc.exe" 
Else 
 Tasks("Calculator").Activate 
End If 
Tasks("Calculator").WindowState = wdWindowStateNormal
```


## See also


#### Concepts


[Tasks Collection Object](tasks-object-word.md)

