---
title: AutoCorrect.ReplaceText Property (Word)
keywords: vbawd10.chm155779076
f1_keywords:
- vbawd10.chm155779076
ms.prod: word
api_name:
- Word.AutoCorrect.ReplaceText
ms.assetid: 4325928d-dc53-4b3c-b6fa-860c090e90e2
ms.date: 06/08/2017
---


# AutoCorrect.ReplaceText Property (Word)

 **True** if Microsoft Word automatically replaces specified text with entries from the AutoCorrect list. Read/write **Boolean** .


## Syntax

 _expression_ . **ReplaceText**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Example

This example sets Word to automatically replace specified text with entries from the AutoCorrect list as you type.


```vb
AutoCorrect.ReplaceText = True
```

This example toggles the value of the ReplaceText property.




```vb
AutoCorrect.ReplaceText = Not AutoCorrect.ReplaceText
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

