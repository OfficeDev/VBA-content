---
title: Options.AutoFormatAsYouTypeApplyBulletedLists Property (Word)
keywords: vbawd10.chm162988294
f1_keywords:
- vbawd10.chm162988294
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeApplyBulletedLists
ms.assetid: 5e077bf3-3db0-a7ab-0bb0-89476b6d3a2c
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeApplyBulletedLists Property (Word)

 **True** if bullet characters (such as asterisks, hyphens, and greater-than signs) are replaced with bullets from the **Bullets And Numbering** dialog box ( **Format** menu) as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyBulletedLists**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


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


[Options Object](options-object-word.md)

