---
title: Options.AutoFormatPreserveStyles Property (Word)
keywords: vbawd10.chm162988291
f1_keywords:
- vbawd10.chm162988291
ms.prod: word
api_name:
- Word.Options.AutoFormatPreserveStyles
ms.assetid: cbde64c7-4a82-f33f-c337-bbc24c32ab40
ms.date: 06/08/2017
---


# Options.AutoFormatPreserveStyles Property (Word)

 **True** if previously applied styles are preserved when Word formats a document or range automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatPreserveStyles**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example sets Word to preserve existing styles and to format headings, lists, and other paragraphs with styles when formatting automatically. Word then formats the current selection automatically.


```vb
With Options 
 .AutoFormatPreserveStyles = True 
 .AutoFormatApplyHeadings = True 
 .AutoFormatApplyLists = True 
 .AutoFormatApplyOtherParas = True 
End With 
Selection.Range.AutoFormat
```

This example returns the status of the  **Styles** option on the **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatPreserveStyles
```


## See also


#### Concepts


[Options Object](options-object-word.md)

