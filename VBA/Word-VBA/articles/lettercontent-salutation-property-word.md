---
title: LetterContent.Salutation Property (Word)
keywords: vbawd10.chm161546350
f1_keywords:
- vbawd10.chm161546350
ms.prod: word
api_name:
- Word.LetterContent.Salutation
ms.assetid: 115a740f-720f-a7d7-df68-148cd36b22c0
ms.date: 06/08/2017
---


# LetterContent.Salutation Property (Word)

Returns or sets the salutation text for a letter created by the Letter Wizard. Read/write  **String** .


## Syntax

 _expression_ . **Salutation**

 _expression_ An expression that returns a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example creates a new  **LetterContent** object, sets several properties (including the salutation text), and then runs the Letter Wizard by using the **[RunLetterWizard](document-runletterwizard-method-word.md)** method.


```vb
Set myContent = New LetterContent 
myContent.Salutation ="Hello," 
Documents.Add.RunLetterWizard LetterContent:=myContent
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

