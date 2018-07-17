---
title: Options.AutoFormatAsYouTypeApplyHeadings Property (Word)
keywords: vbawd10.chm162988292
f1_keywords:
- vbawd10.chm162988292
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeApplyHeadings
ms.assetid: 19dfb55e-8a5c-4e6e-a909-02adcb5a76e9
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeApplyHeadings Property (Word)

 **True** if styles are automatically applied to headings as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyHeadings**

 _expression_ A variable that represents an **[Options](options-object-word.md)** collection.


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


[Options Object](options-object-word.md)

