---
title: Options.AutoFormatApplyLists Property (Word)
keywords: vbawd10.chm162988283
f1_keywords:
- vbawd10.chm162988283
ms.prod: word
api_name:
- Word.Options.AutoFormatApplyLists
ms.assetid: f5d2e1d2-01f8-c3ca-565c-d8cf767741bd
ms.date: 06/08/2017
---


# Options.AutoFormatApplyLists Property (Word)

 **True** if styles are automatically applied to lists when Word formats a document or range automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatApplyLists**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example applies styles to any lists in the current selection.


```vb
Options.AutoFormatApplyLists = True 
Selection.Range.AutoFormat
```

This example returns the status of the  **Lists** option on the **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatApplyLists
```


## See also


#### Concepts


[Options Object](options-object-word.md)

