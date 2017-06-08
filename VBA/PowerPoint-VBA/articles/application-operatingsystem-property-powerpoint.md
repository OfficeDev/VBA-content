---
title: Application.OperatingSystem Property (PowerPoint)
keywords: vbapp10.chm502016
f1_keywords:
- vbapp10.chm502016
ms.prod: powerpoint
api_name:
- PowerPoint.Application.OperatingSystem
ms.assetid: 5532197a-f6c3-825a-6492-e1c85d97a9d2
ms.date: 06/08/2017
---


# Application.OperatingSystem Property (PowerPoint)

Returns the name of the operating system. Read-only.


## Syntax

 _expression_. **OperatingSystem**

 _expression_ A variable that represents an **Application** object.


### Return Value

String


## Example

This example tests the  **OperatingSystem** property to see whether Microsoft PowerPoint is running with a 32-bit version of Microsoft Windows.


```
os = Application.OperatingSystem

If InStr(os, "Windows (32-bit)") <> 0 Then

    MsgBox "Running a 32-bit version of Microsoft Windows"

End If
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

