---
title: Application.Build Property (Publisher)
keywords: vbapb10.chm131078
f1_keywords:
- vbapb10.chm131078
ms.prod: publisher
api_name:
- Publisher.Application.Build
ms.assetid: e0d4bb8e-5185-3d3c-fd80-c1e3c3902b2c
ms.date: 06/08/2017
---


# Application.Build Property (Publisher)

Returns a  **Long** that represents the Microsoft Publisher build number. Read-only.


## Syntax

 _expression_. **Build**

 _expression_A variable that represents a  **Application** object.


### Return Value

Long


## Example

This example displays the Publisher build number.


```vb
Sub BuildNumber() 
 MsgBox Prompt:="The current Microsoft Publisher build number is " &; _ 
 Application.Build, Title:="Microsoft Publisher Build" 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

