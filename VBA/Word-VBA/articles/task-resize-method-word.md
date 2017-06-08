---
title: Task.Resize Method (Word)
keywords: vbawd10.chm159514637
f1_keywords:
- vbawd10.chm159514637
ms.prod: word
api_name:
- Word.Task.Resize
ms.assetid: e4176266-c511-3f4c-f22c-ec5617cd41d9
ms.date: 06/08/2017
---


# Task.Resize Method (Word)

Sizes the specified task window.


## Syntax

 _expression_ . **Resize**( **_Width_** , **_Height_** )

 _expression_ Required. A variable that represents a **[Task](task-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Width_|Required| **Long**|The width of the window, in points.|
| _Height_|Required| **Long**|The height of the window, in points.|

## Remarks

If the window is maximized or minimized, using this method causes an error. Use the  **Width** or **Height** property to set the window width and height independently.


## Example

This example resizes the Microsoft Excel application window to 6 inches wide by 4 inches high.


```vb
If Tasks.Exists("Microsoft Excel") = True Then 
 With Tasks("Microsoft Excel") 
 .WindowState = wdWindowStateNormal 
 .Resize Width:=InchesToPoints(6), Height:=InchesToPoints(4) 
 End With 
End If
```


## See also


#### Concepts


[Task Object](task-object-word.md)

