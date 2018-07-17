---
title: EmailOptions.AutoFormatAsYouTypeReplaceSymbols Property (Word)
keywords: vbawd10.chm165347593
f1_keywords:
- vbawd10.chm165347593
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeReplaceSymbols
ms.assetid: d8314d95-5701-51a7-a987-10cf22f1f87e
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeReplaceSymbols Property (Word)

 **True** if two consecutive hyphens (--) are replaced with an en dash (-) or an em dash (—) as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceSymbols**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Remarks

If the hyphens are typed with leading and trailing spaces, Word replaces the hyphens with an en dash; if there are no trailing spaces, the hyphens are replaced with an em dash.


## Example

This example turns on the replacement of hyphens with symbols as you type.


```vb
EmailOptions.AutoFormatAsYouTypeReplaceSymbols = True
```

This example returns the status of the  **Symbol characters (--) with symbols (—)** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = EmailOptions.AutoFormatAsYouTypeReplaceSymbols
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

