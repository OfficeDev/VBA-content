---
title: EmailOptions.HTMLFidelity Property (Word)
keywords: vbawd10.chm165347635
f1_keywords:
- vbawd10.chm165347635
ms.prod: word
api_name:
- Word.EmailOptions.HTMLFidelity
ms.assetid: 4b9107af-9af5-7691-9237-c3173c71fcd4
ms.date: 06/08/2017
---


# EmailOptions.HTMLFidelity Property (Word)

Strips HTML tags used for opening HTML files in Word but not required for display. Read/write  **WdEmailHTMLFidelity** .


## Syntax

 _expression_ . **HTMLFidelity**

 _expression_ Required. A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example keeps all HTML tags intact when sending e-mail messages.


```vb
Sub HTMLEmail() 
 Application.EmailOptions _ 
 .HTMLFidelity = wdEmailHTMLFidelityHigh 
End Sub
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

