---
title: EmailOptions.PlainTextStyle Property (Word)
keywords: vbawd10.chm165347445
f1_keywords:
- vbawd10.chm165347445
ms.prod: word
api_name:
- Word.EmailOptions.PlainTextStyle
ms.assetid: e3359d77-8ea6-4026-3125-c13436b4e34f
ms.date: 06/08/2017
---


# EmailOptions.PlainTextStyle Property (Word)

Returns the  **[Style](style-object-word.md)** object that represents the text attributes for e-mail messages that are sent or received using plain text.


## Syntax

 _expression_ . **PlainTextStyle**

 _expression_ A variable that represents a **[EmailOptions](emailoptions-object-word.md)** object.


## Example

This example sets the plain text font for e-mail messages to Tahoma, size 10.


```vb
Sub PlainTxt() 
 With Application.EmailOptions.PlainTextStyle 
 .Font.Name = "Tahoma" 
 .Font.Size = 10 
 End With 
End Sub
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

