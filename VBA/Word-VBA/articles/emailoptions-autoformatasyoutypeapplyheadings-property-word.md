---
title: EmailOptions.AutoFormatAsYouTypeApplyHeadings Property (Word)
keywords: vbawd10.chm165347588
f1_keywords:
- vbawd10.chm165347588
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeApplyHeadings
ms.assetid: 299897d1-1132-4ba2-d3e6-47d34a4c38ae
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeApplyHeadings Property (Word)

 **True** if styles are automatically applied to headings as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyHeadings**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example sets Word to automatically apply the Heading1 through Heading 9 styles to headings as you type.


```vb
Options.AutoFormatAsYouTypeApplyHeadings = True
```

This example returns the status of the  **Headings** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeApplyHeadings
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

