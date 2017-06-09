---
title: Application.OnTime Method (Word)
keywords: vbawd10.chm158335326
f1_keywords:
- vbawd10.chm158335326
ms.prod: word
api_name:
- Word.Application.OnTime
ms.assetid: 732d03cc-9dd6-5961-9763-048f72dea4d2
ms.date: 06/08/2017
---


# Application.OnTime Method (Word)

Starts a background timer that runs a macro at a specified time.


## Syntax

 _expression_ . **OnTime**( **_When_** , **_Name_** , **_Tolerance_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _When_|Required| **Variant**|The time at which the macro is to be run. |
| _Name_|Required| **String**|The name of the macro to be run. |
| _Tolerance_|Optional| **Variant**|The maximum time (in seconds) that can elapse before a macro that was not run at the time specified by When is canceled. Macros may not always run at the specified time. For example, if a sort operation is under way or a dialog box is being displayed, the macro will be delayed until Word has completed the task. If this argument is 0 (zero) or omitted, the macro is run regardless of how much time has elapsed since the time specified by When.|

## Remarks

The When parameter can be a string that specifies a time (for example,  `"4:30 pm"` or `"16:30"`), or it can be a serial number returned by a function such as  **TimeValue** or **TimeSerial** (for example, `TimeValue("2:30 pm")` or `TimeSerial(14, 30, 00)`). You can also include the date (for example,  `"6/30 4:15 pm"` or `TimeValue("6/30 4:15 pm")`).

For the Name parameter, use the complete macro path to ensure that the correct macro is run (for example,  `"Project.Module1.Macro1"`). For the macro to run, the document or template must be available both when the  **OnTime** instruction is run and when the time specified by When arrives. For this reason, it is best to store the macro in Normal.dot or another global template that's loaded automatically.

Use the sum of the return values of the  **Now** function and either the **TimeValue** or **TimeSerial** function to set a timer to run a macro a specified amount of time after the statement is run. For example, use `Now+TimeValue("00:05:30")` to run a macro 5 minutes and 30 seconds after the statement is run.

Word can maintain only one background timer set by  **OnTime** . If you start another timer before an existing timer runs, the existing timer is canceled.


## Example

This example runs the macro named "Macro1" in the current module at 3:55 P.M.


```vb
Application.OnTime When:="15:55:00", Name:="Macro1"
```

This example runs the macro named "Macro1" 15 seconds from the time the example is run. The macro name includes the project and module name.




```vb
Application.OnTime When:=Now + TimeValue("00:00:15"), _ 
 Name:="Project1.Module1.Macro1"
```

This example runs the macro named "Start" at 1:30 P.M. The macro name includes the project and module name.




```vb
Application.OnTime When:=TimeValue("1:30 pm"), _ 
 Name:="VBAProj.Module1.Start"
```


## See also


#### Concepts


[Application Object](application-object-word.md)

