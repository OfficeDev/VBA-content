---
title: EmailOptions.AutoFormatAsYouTypeApplyBulletedLists Property (Word)
keywords: vbawd10.chm165347590
f1_keywords:
- vbawd10.chm165347590
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeApplyBulletedLists
ms.assetid: b8bb6d3f-2226-db63-6edd-e8313a13c8c7
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeApplyBulletedLists Property (Word)

 **True** if bullet characters (such as asterisks, hyphens, and greater-than signs) are replaced with bullets. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyBulletedLists**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Remarks

If set to  **True** , Word replaces bullet character with bullets defined in the **Bullets And Numbering** dialog box ( **Format** menu) as you type.


## Example

This example causes characters to be replaced with bullets when typed in a list.


```vb
Options.AutoFormatAsYouTypeApplyBulletedLists = True
```

This example returns the status of the  **Automatic bulleted lists** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeApplyBulletedLists
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

