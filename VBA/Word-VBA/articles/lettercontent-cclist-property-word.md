---
title: LetterContent.CCList Property (Word)
keywords: vbawd10.chm161546358
f1_keywords:
- vbawd10.chm161546358
ms.prod: word
api_name:
- Word.LetterContent.CCList
ms.assetid: 87e4fd7c-ae2e-bb29-c228-65c217a41976
ms.date: 06/08/2017
---


# LetterContent.CCList Property (Word)

Returns or sets the carbon copy (CC) recipients for a letter created by the Letter Wizard. Read/write  **String** .


## Syntax

 _expression_ . **CCList**

 _expression_ A variable that represents a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example displays the CC list text for the active document.


```vb
MsgBox ActiveDocument.GetLetterContent.CCList
```

This example creates a new  **LetterContent** object, sets the courtesy copies by setting the CClist property, and then runs the Letter Wizard by using the **[RunLetterWizard](document-runletterwizard-method-word.md)** method.




```vb
Dim lcNew As New LetterContent 
 
lcNew.CCList = "K. Jordan, D. Funk, D. Morrison" 
ActiveDocument.RunLetterWizard LetterContent:=lcNew
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

