---
title: Application.EmailTemplate Property (Word)
keywords: vbawd10.chm158335427
f1_keywords:
- vbawd10.chm158335427
ms.prod: word
api_name:
- Word.Application.EmailTemplate
ms.assetid: 339a421e-b608-4063-a6e8-a08ba4debaf5
ms.date: 06/08/2017
---


# Application.EmailTemplate Property (Word)

Returns or sets a  **String** that represents the document template to use when sending e-mail messages. Read/write.


## Syntax

 _expression_ . **EmailTemplate**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

Use the  **EmailTemplate** property when Microsoft Word is specified as your e-mail editor, which you must do inside Microsoft Outlook.


## Example

This example instructs Word to use the template named "Email" for all new e-mail messages. This example assumes that you have a template named "Email" and that it is stored in the default template location.


```vb
Sub MessageTemplate() 
 Application.EmailTemplate = "Email" 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

