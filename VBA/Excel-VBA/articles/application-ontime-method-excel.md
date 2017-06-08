---
title: Application.OnTime Method (Excel)
keywords: vbaxl10.chm133184
f1_keywords:
- vbaxl10.chm133184
ms.prod: excel
api_name:
- Excel.Application.OnTime
ms.assetid: 31268da0-8ec7-7169-a1d0-8db34b3385cd
ms.date: 06/08/2017
---


# Application.OnTime Method (Excel)

Schedules a procedure to be run at a specified time in the future (either at a specific time of day or after a specific amount of time has passed).


## Syntax

 _expression_ . **OnTime**( **_EarliestTime_** , **_Procedure_** , **_LatestTime_** , **_Schedule_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EarliestTime_|Required| **Variant**|The time when you want this procedure to be run.|
| _Procedure_|Required| **String**|The name of the procedure to be run.|
| _LatestTime_|Optional| **Variant**|The latest time at which the procedure can be run. For example, if  _LatestTime_ is set to _EarliestTime_ + 30 and Microsoft Excel is not in Ready, Copy, Cut, or Find mode at _EarliestTime_ because another procedure is running, Microsoft Excel will wait 30 seconds for the first procedure to complete. If Microsoft Excel is not in Ready mode within 30 seconds, the procedure won?t be run. If this argument is omitted, Microsoft Excel will wait until the procedure can be run.|
| _Schedule_|Optional| **Variant**| **True** to schedule a new OnTime procedure. **False** to clear a previously set procedure. The default value is **True** .|

## Remarks

Use  `Now + TimeValue(time)` to schedule something to be run when a specific amount of time (counting from now) has elapsed. Use `TimeValue(time)` to schedule something to be run a specific time.

The value of  _EarliestTime_ is rounded to the closest second.

Set  _Schedule_ to **false** to clear a procedure previously set with the same _Procedure_ and _EarliestTime_ values.

_Procedure_ must take no arguments and cannot be declared in a custom class or form.


## Example

This example runs my_Procedure 15 seconds from now.


```vb
Application.OnTime Now + TimeValue("00:00:15"), "my_Procedure"
```

This example runs my_Procedure at 5 P.M.




```vb
Application.OnTime TimeValue("17:00:00"), "my_Procedure"
```

This example cancels the  **OnTime** setting from the previous example.




```vb
Application.OnTime EarliestTime:=TimeValue("17:00:00"), _ 
 Procedure:="my_Procedure", Schedule:=False
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

