---
title: Application.Build Property (Access)
keywords: vbaac10.chm12600
f1_keywords:
- vbaac10.chm12600
ms.prod: access
api_name:
- Access.Application.Build
ms.assetid: d96de996-33f5-a4a1-66d9-c18b3cdbac43
ms.date: 06/08/2017
---


# Application.Build Property (Access)

Returns as a  **Long** representing the build number of the currently installed copy of Microsoft Access. Read-only.


## Syntax

 _expression_. **Build**

 _expression_ A variable that represents an **Application** object.


## Example

The following example displays the version and build number of the currently-installed copy of Microsoft Access.


```vb
MsgBox "You are currently running Microsoft Access, " _ 
 &; " version " &; Application.Version &; ", build " _ 
 &; Application.Build &; "."
```


## See also


#### Concepts


[Application Object](application-object-access.md)

