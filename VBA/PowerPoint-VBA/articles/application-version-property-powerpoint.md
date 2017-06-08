---
title: Application.Version Property (PowerPoint)
keywords: vbapp10.chm502015
f1_keywords:
- vbapp10.chm502015
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Version
ms.assetid: c76b1e7e-db29-0ef8-fefb-9333b8350de0
ms.date: 06/08/2017
---


# Application.Version Property (PowerPoint)

Returns the Microsoft PowerPoint version number. Read-only.


## Syntax

 _expression_. **Version**

 _expression_ A variable that represents a **Application** object.


### Return Value

String


## Example

This example displays a message box that contains the PowerPoint version number and build number, and the name of the operating system.


```vb
With Application
    MsgBox "Welcome to PowerPoint version " &; .Version &; _
        ", build " &; .Build &; ", running on " &; .OperatingSystem &; "!"
End With
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

