---
title: Application.ProductCode Method (Word)
keywords: vbawd10.chm158335380
f1_keywords:
- vbawd10.chm158335380
ms.prod: word
api_name:
- Word.Application.ProductCode
ms.assetid: 3913ee8b-291b-e81c-b106-01007738c7a0
ms.date: 06/08/2017
---


# Application.ProductCode Method (Word)

Returns the Microsoft Word globally unique identifier (GUID) as a  **String** .


## Syntax

 _expression_ . **ProductCode**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Return Value

String


## Example

This example displays the GUID for Microsoft Word.


```vb
MsgBox Application.ProductCode
```


## See also


#### Concepts


[Application Object](application-object-word.md)

