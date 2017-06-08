---
title: Application.EnableCancelKey Property (Excel)
keywords: vbaxl10.chm133131
f1_keywords:
- vbaxl10.chm133131
ms.prod: excel
api_name:
- Excel.Application.EnableCancelKey
ms.assetid: 7c9c17b3-dd04-c914-4ed5-a6ef81ccf0c3
ms.date: 06/08/2017
---


# Application.EnableCancelKey Property (Excel)

Controls how Microsoft Excel handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure. Read/write  **[XlEnableCancelKey](xlenablecancelkey-enumeration-excel.md)** .


## Syntax

 _expression_ . **EnableCancelKey**

 _expression_ A variable that represents an **Application** object.


## Remarks



| **XlEnableCancelKey** can be one of these **XlEnableCancelKey** constants.|
| **xlDisabled** . Cancel key trapping is completely disabled.|
| **xlErrorHandler** . The interrupt is sent to the running procedure as an error, trappable by an error handler set up with an `On Error GoTo` statement. The trappable error code is 18.|
| **xlInterrupt** . The current procedure is interrupted, and the user can debug or end the procedure.|
Use this property very carefully. If you use  **xlDisabled** , there's no way to interrupt a runaway loop or other non - self-terminating code. Likewise, if you use **xlErrorHandler** but your error handler always returns using the `Resume` statement, there's no way to stop runaway code.

The  **EnableCancelKey** property is always reset to **xlInterrupt** whenever Microsoft Excel returns to the idle state and there's no code running. To trap or disable cancellation in your procedure, you must explicitly change the **EnableCancelKey** property every time the procedure is called.


## Example

This example shows how you can use the  **EnableCancelKey** property to set up a custom cancellation handler.


```vb
On Error GoTo handleCancel 
Application.EnableCancelKey = xlErrorHandler 
MsgBox "This may take a long time: press ESC to cancel" 
For x = 1 To 1000000 ' Do something 1,000,000 times (long!) 
 ' do something here 
Next x 
 
handleCancel: 
If Err = 18 Then 
 MsgBox "You cancelled" 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

