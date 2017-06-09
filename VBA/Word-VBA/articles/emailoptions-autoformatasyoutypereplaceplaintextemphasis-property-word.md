---
title: EmailOptions.AutoFormatAsYouTypeReplacePlainTextEmphasis Property (Word)
keywords: vbawd10.chm165347596
f1_keywords:
- vbawd10.chm165347596
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeReplacePlainTextEmphasis
ms.assetid: 2fbd053f-cc0b-e38b-2f2a-dfc0c7f49a38
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeReplacePlainTextEmphasis Property (Word)

 **True** if manual emphasis characters are automatically replaced with character formatting as you type; for example, "*bold*" is changed to " **bold** ". Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplacePlainTextEmphasis**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example turns on the replacement of manual emphasis characters with character formatting.


```vb
Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = True
```

This example returns the status of the  ***Bold* and _underline_ with real formatting** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = _ 
 Options.AutoFormatAsYouTypeReplacePlainTextEmphasis
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

