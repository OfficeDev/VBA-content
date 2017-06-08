---
title: AutoCorrect.CorrectTableCells Property (Word)
keywords: vbawd10.chm155779091
f1_keywords:
- vbawd10.chm155779091
ms.prod: word
api_name:
- Word.AutoCorrect.CorrectTableCells
ms.assetid: 8bb5dfdd-9c54-b49e-609f-18b4d8b556ee
ms.date: 06/08/2017
---


# AutoCorrect.CorrectTableCells Property (Word)

 **True** to automatically capitalize the first letter of table cells. Read/write **Boolean** .


## Syntax

 _expression_ . **CorrectTableCells**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Example

This example disables automatic capitalization of the first letter typed within table cells.


```vb
Sub AutoCorrectFirstLetterOfTableCells() 
 Application.AutoCorrect.CorrectTableCells = False 
End Sub
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

