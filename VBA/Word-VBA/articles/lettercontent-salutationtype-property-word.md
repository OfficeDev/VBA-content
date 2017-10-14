---
title: LetterContent.SalutationType Property (Word)
keywords: vbawd10.chm161546351
f1_keywords:
- vbawd10.chm161546351
ms.prod: word
api_name:
- Word.LetterContent.SalutationType
ms.assetid: f312bdfd-a10d-144d-4b99-0984707d13cb
ms.date: 06/08/2017
---


# LetterContent.SalutationType Property (Word)

Returns or sets the type of salutation for a letter created by the Letter Wizard. Read/write  **WdSalutationType** .


## Syntax

 _expression_ . **SalutationType**

 _expression_ Required. A variable that represents a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example creates a new  **LetterContent** object, sets several properties (including the salutation text), and then runs the Letter Wizard by using the **RunLetterWizard** method.


```vb
Set myContent = New LetterContent 
myContent.SalutationType = wdSalutationBusiness 
Documents.Add.RunLetterWizard _ 
 LetterContent:=myContent, WizardMode:=True
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

