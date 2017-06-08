---
title: Options.AutoFormatReplaceQuotes Property (Word)
keywords: vbawd10.chm162988286
f1_keywords:
- vbawd10.chm162988286
ms.prod: word
api_name:
- Word.Options.AutoFormatReplaceQuotes
ms.assetid: 23fe2823-0aec-7deb-8fc1-ff70a79b19af
ms.date: 06/08/2017
---


# Options.AutoFormatReplaceQuotes Property (Word)

 **True** if straight quotation marks are automatically changed to smart (curly) quotation marks when Word formats a document or range automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatReplaceQuotes**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example turns on the automatic replacement of straight quotation marks with smart (curly) quotation marks, and then it formats the current selection automatically.


```vb
Options.AutoFormatReplaceQuotes = True 
Selection.Range.AutoFormat
```

This example returns the status of the  **Straight quotes with smart quotes** option on the **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatReplaceQuotes
```


## See also


#### Concepts


[Options Object](options-object-word.md)

