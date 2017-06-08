---
title: Application.EnableCancelKey Property (Word)
keywords: vbawd10.chm158335076
f1_keywords:
- vbawd10.chm158335076
ms.prod: word
api_name:
- Word.Application.EnableCancelKey
ms.assetid: dd7d6885-7306-c6f3-56ff-e6f828adc4ea
ms.date: 06/08/2017
---


# Application.EnableCancelKey Property (Word)

Returns or sets the way that Word handles CTRL+BREAK user interruptions. Read/write  **WdEnableCancelKey** .


## Syntax

 _expression_ . **EnableCancelKey**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

Use this property very carefully. If you use  **wdCancelDisabled** , there is no way to interrupt a runaway loop or other non - self-terminating code. Also, the **EnableCancelKey** property is not reset to **wdCancelInterrupt** when your code stops running; unless you explicitly reset its value, it will remain set to **wdCancelDisabled** for the duration of the Word session.


## Example

This example disables CTRL+BREAK from interrupting a counter loop.


```vb
Dim intWait As Integer 
 
Application.EnableCancelKey = wdCancelDisabled 
For intWait = 1 To 10000 
 StatusBar = intWait 
Next intWait 
Application.EnableCancelKey = wdCancelInterrupt
```


## See also


#### Concepts


[Application Object](application-object-word.md)

