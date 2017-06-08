---
title: AutoCorrectEntry.Value Property (Word)
keywords: vbawd10.chm155648003
f1_keywords:
- vbawd10.chm155648003
ms.prod: word
api_name:
- Word.AutoCorrectEntry.Value
ms.assetid: 20744fd3-5c61-4998-a08b-e0062f3589bb
ms.date: 06/08/2017
---


# AutoCorrectEntry.Value Property (Word)

Returns or sets the value of the AutoCorrect entry. Read/write  **String** .


## Syntax

 _expression_ . **Value**

 _expression_ Required. A variable that represents an **[AutoCorrectEntry](autocorrectentry-object-word.md)** object.


## Remarks

The  **Value** property only returns the first 255 characters of the object's value.


## Example

This example creates an AutoCorrect entry and then displays the value of the new entry.


```
AutoCorrect.Entries.Add Name:="i.e.", Value:="that is" 
MsgBox AutoCorrect.Entries("i.e.").Value
```


## See also


#### Concepts


[AutoCorrectEntry Object](autocorrectentry-object-word.md)

