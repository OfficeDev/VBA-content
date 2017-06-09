---
title: Options.AutoFormatReplacePlainTextEmphasis Property (Word)
keywords: vbawd10.chm162988290
f1_keywords:
- vbawd10.chm162988290
ms.prod: word
api_name:
- Word.Options.AutoFormatReplacePlainTextEmphasis
ms.assetid: a01034cc-18b0-425f-8296-884382a17b3c
ms.date: 06/08/2017
---


# Options.AutoFormatReplacePlainTextEmphasis Property (Word)

 **True** if manual emphasis characters are replaced with character formatting when Word formats a document or range automatically. For example, "*bold*" is changed to "bold" and "_underline_" is changed to "underline." Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatReplacePlainTextEmphasis**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example turns on the replacement of manual emphasis characters with character formatting


```vb
Options.AutoFormatReplacePlainTextEmphasis = True 
Selection.Range.AutoFormat
```

This example returns the status of the *Bold* and _underline_ with real formatting option on the  **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatReplacePlainTextEmphasis
```


## See also


#### Concepts


[Options Object](options-object-word.md)

