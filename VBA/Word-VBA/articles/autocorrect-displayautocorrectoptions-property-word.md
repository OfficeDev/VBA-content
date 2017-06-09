---
title: AutoCorrect.DisplayAutoCorrectOptions Property (Word)
keywords: vbawd10.chm155779092
f1_keywords:
- vbawd10.chm155779092
ms.prod: word
api_name:
- Word.AutoCorrect.DisplayAutoCorrectOptions
ms.assetid: 7a4d6773-53f7-8d9d-499e-8d32917c14fd
ms.date: 06/08/2017
---


# AutoCorrect.DisplayAutoCorrectOptions Property (Word)

 **True** for Microsoft Word to display the **AutoCorrect Options** button. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayAutoCorrectOptions**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Example

This example disables display of the  **AutoCorrect Options** button.


```vb
Sub HideAutoCorrectOpButton() 
 AutoCorrect.DisplayAutoCorrectOptions = False 
End Sub
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

