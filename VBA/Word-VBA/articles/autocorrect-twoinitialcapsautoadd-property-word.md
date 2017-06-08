---
title: AutoCorrect.TwoInitialCapsAutoAdd Property (Word)
keywords: vbawd10.chm155779082
f1_keywords:
- vbawd10.chm155779082
ms.prod: word
api_name:
- Word.AutoCorrect.TwoInitialCapsAutoAdd
ms.assetid: 93030da5-453a-392a-3dc4-3c30a12cbea1
ms.date: 06/08/2017
---


# AutoCorrect.TwoInitialCapsAutoAdd Property (Word)

 **True** if Microsoft Word automatically adds words to the list of AutoCorrect Initial Caps exceptions. A word is added to this list if you delete and then retype the uppercase letter (following the initial uppercase letter) that Word changed to lowercase. Read/write **Boolean** .


## Syntax

 _expression_ . **TwoInitialCapsAutoAdd**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Example

This example sets Word to automatically add words to the list of AutoCorrect Initial Caps exceptions.


```vb
AutoCorrect.TwoInitialCapsAutoAdd = True
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

