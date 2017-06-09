---
title: Options.AutoFormatAsYouTypeDefineStyles Property (Word)
keywords: vbawd10.chm162988302
f1_keywords:
- vbawd10.chm162988302
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeDefineStyles
ms.assetid: 16657544-0185-204f-1cee-b959c91956d5
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeDefineStyles Property (Word)

 **True** if Word automatically creates new styles based on manual formatting. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeDefineStyles**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Word to automatically create styles as you type.


```vb
Options.AutoFormatAsYouTypeDefineStyles = True
```

This example returns the status of the Define styles based on your formatting option on the AutoFormat As You Type tab in the AutoCorrect dialog box (Tools menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeDefineStyles
```


## See also


#### Concepts


[Options Object](options-object-word.md)

