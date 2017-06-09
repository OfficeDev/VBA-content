---
title: Application.StartupPath Property (Word)
keywords: vbawd10.chm158335059
f1_keywords:
- vbawd10.chm158335059
ms.prod: word
api_name:
- Word.Application.StartupPath
ms.assetid: 1b73f234-358b-a360-ca69-ed00e0817038
ms.date: 06/08/2017
---


# Application.StartupPath Property (Word)

Returns or sets the complete path of the startup folder, excluding the final separator. Read/write  **String** .


## Syntax

 _expression_ . **StartupPath**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

Templates and add-ins located in the Startup folder are automatically loaded when you start Word.


## Example

This example displays the complete path of the Startup folder.


```vb
MsgBox Application.StartupPath
```

This example enables the user to change the path of the Startup folder.




```vb
x = MsgBox("Do you want to change the startup path?", vbYesNo, _ 
 "Current path = " &; Application.StartupPath) 
If x = vbYes Then 
 newStartup = InputBox("Type a startup path") 
 Application.StartupPath = newStartup 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

