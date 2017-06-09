---
title: Application.Ready Property (Excel)
keywords: vbaxl10.chm133260
f1_keywords:
- vbaxl10.chm133260
ms.prod: excel
api_name:
- Excel.Application.Ready
ms.assetid: 4b9577ee-0f0c-dd0b-c1dd-90cde2c5fb1e
ms.date: 06/08/2017
---


# Application.Ready Property (Excel)

Returns  **True** when the Microsoft Excel application is ready; **False** when the Excel application is not ready. Read-only **Boolean** .


## Syntax

 _expression_ . **Ready**

 _expression_ A variable that represents an **Application** object.


## Example

In this example, Microsoft Excel checks to see if the  **Ready** property is set to **True** , and if so, a message displays "Application is ready." Otherwise, Excel displays the message "Application is not ready."


```vb
Sub UseReady() 
 
 If Application.Ready = True Then 
 MsgBox "Application is ready." 
 Else 
 MsgBox "Application is not ready." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

