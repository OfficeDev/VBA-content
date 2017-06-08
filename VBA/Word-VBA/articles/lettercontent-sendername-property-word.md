---
title: LetterContent.SenderName Property (Word)
keywords: vbawd10.chm161546360
f1_keywords:
- vbawd10.chm161546360
ms.prod: word
api_name:
- Word.LetterContent.SenderName
ms.assetid: 3f6825d1-98ab-0539-d09b-508697dbfe14
ms.date: 06/08/2017
---


# LetterContent.SenderName Property (Word)

Returns or sets the name of the person creating a letter with the Letter Wizard. Read/write  **String** .


## Syntax

 _expression_ . **SenderName**

 _expression_ An expression that returns a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example creates a new  **LetterContent** object, with the sender name and initials from the **User Information** tab in the **Options** dialog box ( **Tools** menu). The example creates a new document and then runs the Letter Wizard by using the **[RunLetterWizard](document-runletterwizard-method-word.md)** method.


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

