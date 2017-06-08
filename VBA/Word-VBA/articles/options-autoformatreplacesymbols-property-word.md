---
title: Options.AutoFormatReplaceSymbols Property (Word)
keywords: vbawd10.chm162988287
f1_keywords:
- vbawd10.chm162988287
ms.prod: word
api_name:
- Word.Options.AutoFormatReplaceSymbols
ms.assetid: 58a1c811-2fd8-92a9-1f85-6d9beb4223ef
ms.date: 06/08/2017
---


# Options.AutoFormatReplaceSymbols Property (Word)

 **True** if two consecutive hyphens (--) are replaced by an en dash (-) or an em dash (—) when Word formats a document or range automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatReplaceSymbols**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example turns on the replacement of hyphens with symbols, and then it formats the current selection automatically.


```vb
Options.AutoFormatReplaceSymbols = True 
Selection.Range.AutoFormat
```

This example returns the status of the Symbol characters (--) with symbols (—) option on the  **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatReplaceSymbols
```


## See also


#### Concepts


[Options Object](options-object-word.md)

