---
title: Options.AutoFormatApplyBulletedLists Property (Word)
keywords: vbawd10.chm162988284
f1_keywords:
- vbawd10.chm162988284
ms.prod: word
api_name:
- Word.Options.AutoFormatApplyBulletedLists
ms.assetid: a66aacd6-0709-d4ac-0af4-314a386ee39c
ms.date: 06/08/2017
---


# Options.AutoFormatApplyBulletedLists Property (Word)

 **True** if characters (such as asterisks, hyphens, and greater-than signs) at the beginning of list paragraphs are replaced with bullets from the **Bullets and Numbering** dialog box ( **Format** menu) when Word formats a document or range automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatApplyBulletedLists**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example replaces any characters used at the beginning of list paragraphs in the current selection with bullets.


```vb
Options.AutoFormatApplyBulletedLists = True 
Selection.Range.AutoFormat
```

This example returns the status of the  **Automatic bulleted lists** option on the **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatApplyBulletedLists
```


## See also


#### Concepts


[Options Object](options-object-word.md)

