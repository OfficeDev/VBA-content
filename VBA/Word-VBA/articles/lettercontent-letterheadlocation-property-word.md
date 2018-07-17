---
title: LetterContent.LetterheadLocation Property (Word)
keywords: vbawd10.chm161546346
f1_keywords:
- vbawd10.chm161546346
ms.prod: word
api_name:
- Word.LetterContent.LetterheadLocation
ms.assetid: 5e8271fa-23bc-fcf5-ca5c-9139120711e4
ms.date: 06/08/2017
---


# LetterContent.LetterheadLocation Property (Word)

Returns or sets the location of the preprinted letterhead in a letter created by the Letter Wizard. Read/write  **WdLetterheadLocation** .


## Syntax

 _expression_ . **LetterheadLocation**

 _expression_ Required. A variable that represents a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example creates a new  **LetterContent** object, reserves an inch of space at the top of the page for a preprinted letterhead, and then runs the Letter Wizard by using the **RunLetterWizard** method.


```vb
Dim lcNew As LetterContent 
 
Set lcNew = New LetterContent 
 
With lcNew 
 .Letterhead = True 
 .LetterheadLocation = wdLetterTop 
 .LetterheadSize = InchesToPoints(1) 
End With 
 
ActiveDocument.RunLetterWizard LetterContent:=lcNew
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

