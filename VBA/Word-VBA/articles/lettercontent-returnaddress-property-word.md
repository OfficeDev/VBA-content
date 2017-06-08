---
title: LetterContent.ReturnAddress Property (Word)
keywords: vbawd10.chm161546359
f1_keywords:
- vbawd10.chm161546359
ms.prod: word
api_name:
- Word.LetterContent.ReturnAddress
ms.assetid: 6a9bb308-c447-b4e6-1ab9-6f73b29bee12
ms.date: 06/08/2017
---


# LetterContent.ReturnAddress Property (Word)

Returns or sets the return address for a letter created with the Letter Wizard. Read/write  **String** .


## Syntax

 _expression_ . **ReturnAddress**

 _expression_ Required. A variable that represents a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example creates a new  **LetterContent** object, sets the return address and several other properties, and then runs the Letter Wizard by using the **RunLetterWizard** method.


```vb
Dim oLC as New LetterContent 
With oLC 
 .LetterStyle = wdFullBlock 
 .Salutation ="Hello" 
 .SalutationType = wdSalutationOther 
 .ReturnAddress = Application.UserAddress 
End With 
Documents.Add.RunLetterWizard LetterContent:=oLC
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

