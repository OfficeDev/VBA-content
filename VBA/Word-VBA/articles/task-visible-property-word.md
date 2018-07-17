---
title: Task.Visible Property (Word)
keywords: vbawd10.chm159514630
f1_keywords:
- vbawd10.chm159514630
ms.prod: word
api_name:
- Word.Task.Visible
ms.assetid: cc1bb50d-c49d-9230-83ad-940c53c89220
ms.date: 06/08/2017
---


# Task.Visible Property (Word)

 **True** if the specified object is visible. Read/write **Boolean** .


## Syntax

 _expression_ . **Visible**

 _expression_ Required. A variable that represents a **[Task](task-object-word.md)** object.


## Remarks

For any object, some methods and properties may be unavailable if the  **Visible** property is **False** .


## Example

This example hides the Calculator, if it is running. If it is not running, a message is displayed.


```vb
If Tasks.Exists("Calculator") Then 
 Tasks("Calculator").Visible = False 
Else 
 Msgbox "Calculator is not running." 
End If
```


## See also


#### Concepts


[Task Object](task-object-word.md)

