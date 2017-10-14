---
title: Options.AutoFormatAsYouTypeApplyNumberedLists Property (Word)
keywords: vbawd10.chm162988295
f1_keywords:
- vbawd10.chm162988295
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeApplyNumberedLists
ms.assetid: a20be170-7297-0f55-4650-04540fc6f4f9
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeApplyNumberedLists Property (Word)

 **True** if paragraphs are automatically formatted as numbered lists with a numbering scheme from the **Bullets and Numbering** dialog box ( **Format** menu), according to what's typed. For example, if a paragraph starts with "1.1" and a tab character, Word automatically inserts "1.2" and a tab character after the ENTER key is pressed. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyNumberedLists**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example causes lists to be automatically numbered as you type.


```vb
Options.AutoFormatAsYouTypeApplyNumberedLists = True
```

This example returns the status of the Automatic numbered lists option on the AutoFormat As You Type tab in the AutoCorrect dialog box (Tools menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeApplyNumberedLists
```


## See also


#### Concepts


[Options Object](options-object-word.md)

