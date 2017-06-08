---
title: Application.ScreenUpdating Property (Excel)
keywords: vbaxl10.chm133205
f1_keywords:
- vbaxl10.chm133205
ms.prod: excel
api_name:
- Excel.Application.ScreenUpdating
ms.assetid: 08fa0272-faeb-f8f2-c0f2-e001620cc838
ms.date: 06/08/2017
---


# Application.ScreenUpdating Property (Excel)

 **True** if screen updating is turned on. Read/write **Boolean** .


## Syntax

 _expression_ . **ScreenUpdating**

 _expression_ A variable that represents an **Application** object.


## Remarks

Turn screen updating off to speed up your macro code. You won't be able to see what the macro is doing, but it will run faster.

Remember to set the  **ScreenUpdating** property back to **True** when your macro ends.


## Example

This example demonstrates how turning off screen updating can make your code run faster. The example hides every other column on Sheet1, while keeping track of the time it takes to do so. The first time the example hides the columns, screen updating is turned on; the second time, screen updating is turned off. When you run this example, you can compare the respective running times, which are displayed in the message box.


```vb
Dim elapsedTime(2) 
Application.ScreenUpdating = True 
For i = 1 To 2 
 If i = 2 Then Application.ScreenUpdating = False 
 startTime = Time 
 Worksheets("Sheet1").Activate 
 For Each c In ActiveSheet.Columns 
 If c.Column Mod 2 = 0 Then 
 c.Hidden = True 
 End If 
 Next c 
 stopTime = Time 
 elapsedTime(i) = (stopTime - startTime) * 24 * 60 * 60 
Next i 
Application.ScreenUpdating = True 
MsgBox "Elapsed time, screen updating on: " &; elapsedTime(1) &; _ 
 " sec." &; Chr(13) &; _ 
 "Elapsed time, screen updating off: " &; elapsedTime(2) &; _ 
 " sec."
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

