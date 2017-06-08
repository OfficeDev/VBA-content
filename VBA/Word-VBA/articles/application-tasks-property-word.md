---
title: Application.Tasks Property (Word)
keywords: vbawd10.chm158335004
f1_keywords:
- vbawd10.chm158335004
ms.prod: word
api_name:
- Word.Application.Tasks
ms.assetid: e78150fd-8c79-948b-7a46-5e560194aa48
ms.date: 06/08/2017
---


# Application.Tasks Property (Word)

Returns a  **[Tasks](tasks-object-word.md)** collection that represents all the applications that are running.


## Syntax

 _expression_ . **Tasks**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the calculator. If the calculator is not already running, then Word starts the task and then displays the calculator.


```vb
If Tasks.Exists("Calculator") Then 
 With Tasks("Calculator") 
 .Activate 
 .WindowState = wdWindowStateNormal 
 End With 
Else 
 Shell "calc.exe" 
 Tasks("Calculator").WindowState = wdWindowStateNormal 
End If
```

This example checks to see whether Microsoft Excel is currently running. If the task is running, the example activates Microsoft Excel; otherwise, a message box is displayed.




```vb
If Tasks.Exists("Microsoft Excel") = True Then 
 With Tasks("Microsoft Excel") 
 .Activate 
 .WindowState = wdWindowStateMaximize 
 End With 
Else 
 Msgbox "Microsoft Excel is not currently running." 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

