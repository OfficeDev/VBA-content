---
title: Application.Version Property (Publisher)
keywords: vbapb10.chm131121
f1_keywords:
- vbapb10.chm131121
ms.prod: publisher
api_name:
- Publisher.Application.Version
ms.assetid: ffec5bca-cd81-77c6-d80b-e629abfa6dec
ms.date: 06/08/2017
---


# Application.Version Property (Publisher)

Returns a  **String** indicating the version number of the currently-installed copy of Microsoft Publisher. Read-only.


## Syntax

 _expression_. **Version**

 _expression_A variable that represents a  **Application** object.


### Return Value

String


## Example

The following example displays the version and build number of the currently-installed copy of Publisher.


```vb
MsgBox "You are currently running Microsoft Publisher, " _ 
 &; " version " &; Application.Version &; ", build " _ 
 &; Application.Build &; "." 

```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

