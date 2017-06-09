---
title: EmailOptions.AutoFormatAsYouTypeApplyTables Property (Word)
keywords: vbawd10.chm165347618
f1_keywords:
- vbawd10.chm165347618
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeApplyTables
ms.assetid: e7435efc-b4a1-97a3-a7b1-d6e1fabfd0c2
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeApplyTables Property (Word)

 **True** if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyTables**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Remarks

The plus signs become the column borders, and the hyphens become the column widths. 


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


[EmailOptions Object](emailoptions-object-word.md)

