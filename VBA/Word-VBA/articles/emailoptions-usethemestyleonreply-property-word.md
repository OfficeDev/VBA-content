---
title: EmailOptions.UseThemeStyleOnReply Property (Word)
keywords: vbawd10.chm165347446
f1_keywords:
- vbawd10.chm165347446
ms.prod: word
api_name:
- Word.EmailOptions.UseThemeStyleOnReply
ms.assetid: 0d194a90-4977-bae3-29dc-2f69a7d40395
ms.date: 06/08/2017
---


# EmailOptions.UseThemeStyleOnReply Property (Word)

 **True** for Microsoft Word to use a theme when replying to e-mail. Read/write **Boolean** .


## Syntax

 _expression_ . **UseThemeStyleOnReply**

 _expression_ An expression that returns an **[EmailOptions](emailoptions-object-word.md)** object.


## Example

This example tells Word to use a theme when replying to e-mail if Word uses a theme for new messages.


```vb
Sub NewTheme() 
 With Application.EmailOptions 
 If .UseThemeStyle = True Then 
 .UseThemeStyleOnReply = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

