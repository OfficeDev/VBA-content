---
title: AutoRecover.Time Property (Excel)
keywords: vbaxl10.chm696074
f1_keywords:
- vbaxl10.chm696074
ms.prod: excel
api_name:
- Excel.AutoRecover.Time
ms.assetid: 096783b6-77ae-75eb-08cc-fa3978aa6121
ms.date: 06/08/2017
---


# AutoRecover.Time Property (Excel)

Sets or returns the time interval for the  **AutoRecover** object. Permissible values are integers from 1 to 120 minutes. The default value is 10 minutes. Read/write **Long** .


## Syntax

 _expression_ . **Time**

 _expression_ A variable that represents an **AutoRecover** object.


## Remarks

Entering a decimal value will round to the nearest whole number. For example, entering a value of 5.5 is the equivalent of 6.

If time values outside the valid range are entered, Microsoft Excel will revert to the previous time value used.


## Example

The following example sets the AutoRecover time interval to 5 minutes and notifies the user.


```vb
Sub SetTimeValue() 
 
 Application.AutoRecover.Time = 5 
 MsgBox "The AutoRecover time interval is set at " &; _ 
 Application.AutoRecover.Time &; " minutes." 
 
End Sub
```


## See also


#### Concepts


[AutoRecover Object](autorecover-object-excel.md)

