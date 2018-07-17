---
title: LetterContent.Letterhead Property (Word)
keywords: vbawd10.chm161546345
f1_keywords:
- vbawd10.chm161546345
ms.prod: word
api_name:
- Word.LetterContent.Letterhead
ms.assetid: afd847ed-46b2-2539-a4b4-550094974614
ms.date: 06/08/2017
---


# LetterContent.Letterhead Property (Word)

 **True** if space is reserved for a preprinted letterhead in a letter created by the Letter Wizard. Read/write **Boolean** . The **[LetterheadSize](lettercontent-letterheadsize-property-word.md)** property controls the size of the reserved letterhead space.


## Syntax

 _expression_ . **Letterhead**

 _expression_ An expression that returns a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example creates a new  **LetterContent** object, reserves an inch of space at the top of the page for a preprinted letterhead, and then runs the Letter Wizard by using the **[RunLetterWizard](document-runletterwizard-method-word.md)** method.


```vb
Dim lcNew As LetterContent 
 
Set lcNew = New LetterContent 
 
With lcNew 
 .Letterhead = True 
 .LetterheadLocation = wdLetterTop 
 .LetterheadSize = InchesToPoints(1) 
End With 
ActiveDocument.RunLetterWizard _ 
 LetterContent:=lcNew, WizardMode:=True
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

