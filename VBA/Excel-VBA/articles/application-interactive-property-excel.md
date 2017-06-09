---
title: Application.Interactive Property (Excel)
keywords: vbaxl10.chm133150
f1_keywords:
- vbaxl10.chm133150
ms.prod: excel
api_name:
- Excel.Application.Interactive
ms.assetid: fe69429e-8715-770c-3e7a-c06a10a8e850
ms.date: 06/08/2017
---


# Application.Interactive Property (Excel)

 **True** if Microsoft Excel is in interactive mode; this property is usually **True** . If you set the this property to **False** , Microsoft Excel will block all input from the keyboard and mouse (except input to dialog boxes that are displayed by your code). Read/write **Boolean** .


## Syntax

 _expression_ . **Interactive**

 _expression_ A variable that represents an **Application** object.


## Remarks

Blocking user input will prevent the user from interfering with the macro as it moves or activates Microsoft Excel objects.

This property is useful if you're using DDE or OLE Automation to communicate with Microsoft Excel from another application.

If you set this property to  **False** , don't forget to set it back to **True** . Microsoft Excel won't automatically set this property back to **True** when your macro stops running.


## Example

This example sets the  **Interactive** property to **False** while it's using DDE in Windows and then sets this property back to **True** when it's finished. This prevents the user from interfering with the macro.


```vb
Application.Interactive = False 
Application.DisplayAlerts = False 
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\FORMLETR.DOC") 
Application.DDEExecute channelNumber, "[FILEPRINT]" 
Application.DDETerminate channelNumber 
Application.DisplayAlerts = True 
Application.Interactive = True
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

