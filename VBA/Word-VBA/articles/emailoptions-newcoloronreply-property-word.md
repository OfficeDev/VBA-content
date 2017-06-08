---
title: EmailOptions.NewColorOnReply Property (Word)
keywords: vbawd10.chm165347444
f1_keywords:
- vbawd10.chm165347444
ms.prod: word
api_name:
- Word.EmailOptions.NewColorOnReply
ms.assetid: f7878b23-46a3-7950-7b45-28810de58f91
ms.date: 06/08/2017
---


# EmailOptions.NewColorOnReply Property (Word)

 **True** specifies whether a user needs to choose a new color for reply text when replying to e-mail. Read/write **Boolean** .


## Syntax

 _expression_ . **NewColorOnReply**

 _expression_ An expression that returns an **[EmailOptions](emailoptions-object-word.md)** object.


## Remarks

Use the  **NewColorOnReply** property if you want the reply text of e-mail messages sent from Microsoft Word to be a different color than the original message.


## Example

This example checks to see if a user needs to choose a new color for e-mail reply text and, if not, sets the reply font color to blue.


```vb
Sub NewColor() 
 With Application.EmailOptions 
 If .NewColorOnReply = False Then 
 .ReplyStyle.Font.Color = wdColorBlue 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

