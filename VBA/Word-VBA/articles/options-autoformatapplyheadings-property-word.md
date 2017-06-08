---
title: Options.AutoFormatApplyHeadings Property (Word)
keywords: vbawd10.chm162988282
f1_keywords:
- vbawd10.chm162988282
ms.prod: word
api_name:
- Word.Options.AutoFormatApplyHeadings
ms.assetid: 9b1d8936-f6f2-4f01-8583-b9a43a00438b
ms.date: 06/08/2017
---


# Options.AutoFormatApplyHeadings Property (Word)

 **True** if styles are automatically applied to headings when Word formats a document or range automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatApplyHeadings**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example applies the Heading 1 through Heading 9 styles to headings in the current selection.


```vb
Options.AutoFormatApplyHeadings = True 
Selection.Range.AutoFormat
```

This example returns the status of the  **Headings** option on the **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatApplyHeadings
```


## See also


#### Concepts


[Options Object](options-object-word.md)

