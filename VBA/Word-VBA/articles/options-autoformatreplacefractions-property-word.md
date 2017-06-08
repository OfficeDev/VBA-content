---
title: Options.AutoFormatReplaceFractions Property (Word)
keywords: vbawd10.chm162988289
f1_keywords:
- vbawd10.chm162988289
ms.prod: word
api_name:
- Word.Options.AutoFormatReplaceFractions
ms.assetid: e6ee4446-6ec0-766d-cb73-1fdbdb755118
ms.date: 06/08/2017
---


# Options.AutoFormatReplaceFractions Property (Word)

 **True** if typed fractions are replaced with fractions from the current character set when Word formats a document or range automatically. For example, "1/2" is replaced with "½." Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatReplaceFractions**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example turns on the replacement of typed fractions, and thenit formats the current selection automatically.


```vb
Options.AutoFormatReplaceFractions = True 
Selection.Range.AutoFormat
```

This example returns the status of the  **Fractions (1/2) with fraction character (½)** option on the **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatReplaceFractions
```


## See also


#### Concepts


[Options Object](options-object-word.md)

