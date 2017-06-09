---
title: Options.AutoFormatAsYouTypeApplyTables Property (Word)
keywords: vbawd10.chm162988322
f1_keywords:
- vbawd10.chm162988322
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeApplyTables
ms.assetid: 171da8ca-5754-b5fb-12b2-1fcb1461a8fd
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeApplyTables Property (Word)

 **True** if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. The plus signs become the column borders, and the hyphens become the column widths. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyTables**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Word to automatically create tables as you type.


```vb
Options.AutoFormatAsYouTypeApplyTables = True
```

This example returns the status of the  **Tables** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeApplyTables
```


## See also


#### Concepts


[Options Object](options-object-word.md)

