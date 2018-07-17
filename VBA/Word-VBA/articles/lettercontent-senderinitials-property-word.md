---
title: LetterContent.SenderInitials Property (Word)
keywords: vbawd10.chm161546364
f1_keywords:
- vbawd10.chm161546364
ms.prod: word
api_name:
- Word.LetterContent.SenderInitials
ms.assetid: 8c2bdb64-840f-c442-a7b6-28c756701c30
ms.date: 06/08/2017
---


# LetterContent.SenderInitials Property (Word)

Returns or sets the initials of the person creating a letter with the Letter Wizard. Read/write  **String** .


## Syntax

 _expression_ . **SenderInitials**

 _expression_ An expression that returns a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example creates a new  **LetterContent** object with the sender name and initials from the **User Information** tab in the **Options** dialog box ( **Tools** menu). The example creates a new document and then runs the Letter Wizard by using the **[RunLetterWizard](document-runletterwizard-method-word.md)** method.


```vb
Set myContent = New LetterContent 
With myContent 
 .SenderName = Application.UserName 
 .SenderInitials =Application.UserInitials 
End With 
Documents.Add.RunLetterWizard _ 
 LetterContent:=myContent, WizardMode:=True
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

