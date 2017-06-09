---
title: EmailOptions.AutoFormatAsYouTypeDefineStyles Property (Word)
keywords: vbawd10.chm165347598
f1_keywords:
- vbawd10.chm165347598
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeDefineStyles
ms.assetid: ec9df413-17f5-a2c2-4386-7b1d44328b78
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeDefineStyles Property (Word)

 **True** if Word automatically creates new styles based on manual formatting. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeDefineStyles**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example sets Word to automatically create styles as you type.


```vb
Options.AutoFormatAsYouTypeDefineStyles = True
```

This example returns the status of the  **Define styles based on your formatting** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeDefineStyles
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

