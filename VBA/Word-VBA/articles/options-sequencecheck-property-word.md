---
title: Options.SequenceCheck Property (Word)
keywords: vbawd10.chm162988458
f1_keywords:
- vbawd10.chm162988458
ms.prod: word
api_name:
- Word.Options.SequenceCheck
ms.assetid: c2279189-a0ab-15bb-5c8d-00f13075b59a
ms.date: 06/08/2017
---


# Options.SequenceCheck Property (Word)

 **True** to check the sequence of independent characters for South Asian text. Read/write **Boolean** .


## Syntax

 _expression_ . **SequenceCheck**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example enables sequence checking, allowing the user to type a valid sequence of independent characters to form valid character cells in South Asian text.


```vb
Sub CheckSequence() 
 Options.SequenceCheck = True 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

