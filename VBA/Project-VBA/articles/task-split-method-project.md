---
title: Task.Split Method (Project)
ms.prod: project-server
api_name:
- Project.Task.Split
ms.assetid: 847c5cfd-a10f-ea6a-aa49-2e2e88d1840e
ms.date: 06/08/2017
---


# Task.Split Method (Project)

Splits the task into two portions.


## Syntax

 _expression_. **Split**( ** _StartSplitOn_**, ** _EndSplitOn_** )

 _expression_ A variable that represents a **Task** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StartSplitOn_|Required|**Variant**|The start date of the task split. If a time is not specified, the project's default end time for the working period is used.|
| _EndSplitOn_|Required|**Variant**|The end date of the task split. If a time is not specified, the project's default start time for the working period is used. If  _EndSplitOn_ is on or before the date specified with _StartSplitOn_, the split is not created.|

## Example

The following example creates a split in the specified task.


```vb
Sub CreateSplit() 
    Dim WhichTask As Long 
    Dim SplitFrom As Variant, SplitTo As Variant 
 
    WhichTask = InputBox("Enter the ID of the task you would like to split:") 
    SplitFrom = InputBox("Enter the date and time for the start of the" &; _
        " split: " &; vbCrLf &; vbCrLf &; "(The default time is the end" &; _
    " time of the preceding working period.)") 
    SplitTo = InputBox("Enter the date and time for the end of the split:" &; _
        vbCrLf &; vbCrLf &; "(The default time is the start time of the next" &; _
        " working period.)") 
 
    ActiveProject.Tasks(WhichTask).Split SplitFrom, SplitTo 
End Sub
```


