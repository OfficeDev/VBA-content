---
title: EmailOptions.AutoFormatAsYouTypeReplaceHyperlinks Property (Word)
keywords: vbawd10.chm165347600
f1_keywords:
- vbawd10.chm165347600
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeReplaceHyperlinks
ms.assetid: 902775b4-f89e-f5bd-879b-6dd3fe6f2d06
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeReplaceHyperlinks Property (Word)

 **True** if e-mail addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are automatically changed to hyperlinks as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceHyperlinks**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Remarks

Word changes any text that looks like an e-mail address, UNC, or URL to a hyperlink. However, Word does not check the validity of the hyperlink.


## Example

This example enables Word to automatically replace any Internet or network paths with hyperlinks when the paths are typed.


```vb
Options.AutoFormatAsYouTypeReplaceHyperlinks = True
```

This example returns the status of the  **Internet and network paths with hyperlinks** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeReplaceHyperlinks
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

