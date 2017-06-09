---
title: EmailOptions.AutoFormatAsYouTypeApplyNumberedLists Property (Word)
keywords: vbawd10.chm165347591
f1_keywords:
- vbawd10.chm165347591
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeApplyNumberedLists
ms.assetid: 39e50b47-1e1c-4ed8-197c-b99476423187
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeApplyNumberedLists Property (Word)

 **True** if paragraphs are automatically formatted as numbered lists. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyNumberedLists**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Remarks

 If set to **True** , numbered lists use a numbering scheme from the **Bullets and Numbering** dialog box ( **Format** menu), according to what's typed. For example, if a paragraph starts with "1.1" and a tab character, Word automatically inserts "1.2" and a tab character when the ENTER key is pressed.


## Example

This example causes lists to be automatically numbered as you type.


```vb
Options.AutoFormatAsYouTypeApplyNumberedLists = True
```

This example returns the status of the  **Automatic numbered lists** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeApplyNumberedLists
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

