---
title: AutoCorrect.CorrectCapsLock Property (Word)
keywords: vbawd10.chm155779083
f1_keywords:
- vbawd10.chm155779083
ms.prod: word
api_name:
- Word.AutoCorrect.CorrectCapsLock
ms.assetid: 2bbc35cc-3eb3-dc1d-250d-8d4c2a5f9cd3
ms.date: 06/08/2017
---


# AutoCorrect.CorrectCapsLock Property (Word)

 **True** if Word automatically corrects instances in which you use the CAPS LOCK key inadvertently as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **CorrectCapsLock**

 _expression_ A variable that represents an **[AutoCorrect](autocorrect-object-word.md)** object.


## Example

This example determines whether Word is set to automatically correct CAPS LOCK key errors.


```vb
If AutoCorrect.CorrectCapsLock = True Then 
 MsgBox "Correct CAPS LOCK is active." 
Else 
 MsgBox "Correct CAPS LOCK is not active." 
End If
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

