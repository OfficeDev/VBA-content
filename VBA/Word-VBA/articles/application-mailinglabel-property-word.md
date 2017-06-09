---
title: Application.MailingLabel Property (Word)
keywords: vbawd10.chm158334994
f1_keywords:
- vbawd10.chm158334994
ms.prod: word
api_name:
- Word.Application.MailingLabel
ms.assetid: 7eba3273-4a4c-6cdf-004a-4a0d214d6127
ms.date: 06/08/2017
---


# Application.MailingLabel Property (Word)

Returns a  **[MailingLabel](mailinglabel-object-word.md)** object that represents a mailing label.


## Syntax

 _expression_ . **MailingLabel**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

This example creates a new Avery 2160 mini-label document for a specified address.


```
addr = "Dave Edson" &; vbCr &; "123 Skye St." _ 
 &; vbCr &; "Our Town, WA 98004" 
Application.MailingLabel.CreateNewDocument _ 
 Name:="2160 mini", Address:=addr, ExtractAddress:=False
```


## See also


#### Concepts


[Application Object](application-object-word.md)

