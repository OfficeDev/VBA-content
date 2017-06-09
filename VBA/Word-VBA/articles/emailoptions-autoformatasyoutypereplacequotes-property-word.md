---
title: EmailOptions.AutoFormatAsYouTypeReplaceQuotes Property (Word)
keywords: vbawd10.chm165347592
f1_keywords:
- vbawd10.chm165347592
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeReplaceQuotes
ms.assetid: 34be4286-4d36-a338-f103-667d7b8b34a0
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeReplaceQuotes Property (Word)

 **True** if straight quotation marks are automatically changed to smart (curly) quotation marks as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceQuotes**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example turns on the automatic replacement of straight quotation marks with smart (curly) quotation marks as you type.


```vb
Options.AutoFormatAsYouTypeReplaceQuotes = True
```

This example returns the status of the  **Straight quotes with smart quotes** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatReplaceQuotes
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

