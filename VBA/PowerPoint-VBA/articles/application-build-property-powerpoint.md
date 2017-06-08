---
title: Application.Build Property (PowerPoint)
keywords: vbapp10.chm502014
f1_keywords:
- vbapp10.chm502014
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Build
ms.assetid: e485e2f1-835c-33aa-c585-32fbd3af4a88
ms.date: 06/08/2017
---


# Application.Build Property (PowerPoint)

Returns the build number for the current instance of Microsoft PowerPoint. Read-only.


## Syntax

 _expression_. **Build**

 _expression_ A variable that represents a **Application** object.


### Return Value

String


## Example

This example displays the PowerPoint build number.


```vb
MsgBox Prompt:=Application.Build, Title:="PowerPoint Build"
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

