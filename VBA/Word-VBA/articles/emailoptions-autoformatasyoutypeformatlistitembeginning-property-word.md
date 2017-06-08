---
title: EmailOptions.AutoFormatAsYouTypeFormatListItemBeginning Property (Word)
keywords: vbawd10.chm165347597
f1_keywords:
- vbawd10.chm165347597
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeFormatListItemBeginning
ms.assetid: b6450b00-f073-a7f3-2ce4-6fc057a17d41
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeFormatListItemBeginning Property (Word)

 **True** if Word repeats character formatting applied to the beginning of a list item to the next list item. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeFormatListItemBeginning**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example sets Word to automatically repeat character formatting at the beginning of list items.


```vb
Options.AutoFormatAsYouTypeFormatListItemBeginning = True
```

This example returns the status of the  **Format beginning of list item like the one before it** option in the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Options** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = _ 
 Options.AutoFormatAsYouTypeFormatListItemBeginning
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

