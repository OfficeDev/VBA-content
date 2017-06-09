---
title: Options.ContextualSpeller Property (Word)
keywords: vbawd10.chm162988517
f1_keywords:
- vbawd10.chm162988517
ms.prod: word
api_name:
- Word.Options.ContextualSpeller
ms.assetid: d75fc899-5b4e-b30c-661d-4fa2fad0ea38
ms.date: 06/08/2017
---


# Options.ContextualSpeller Property (Word)

Returns or sets a  **Boolean** that represents whether to use the contextual speller to check spelling based on the context of a word and the words around it. Read/write.


## Syntax

 _expression_ . **ContextualSpeller**

 _expression_ An expression that returns an **Options** object.


## Remarks

The contextual speller indentifies the structure of words and their location within a sentence to determine if spelling is correct. It can find errors that the standard spelling checker cannot find. For example, a user might type "superb owl" instead of "super bowl". Because both "superb" and "owl" are correctly spelled words, the standard spelling checker would not find an error. Based on the words in context of the sentence and the words around them, the contextual speller determines that the correct words are "super" and "bowl" and makes the change automatically.

This property corresponds to the  **Use contextual spelling** check box in the **Proofing** tab of the **Word Options** dialog box.


## See also


#### Concepts


[Options Object](options-object-word.md)

