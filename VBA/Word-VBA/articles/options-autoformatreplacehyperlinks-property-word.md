---
title: Options.AutoFormatReplaceHyperlinks Property (Word)
keywords: vbawd10.chm162988305
f1_keywords:
- vbawd10.chm162988305
ms.prod: word
api_name:
- Word.Options.AutoFormatReplaceHyperlinks
ms.assetid: affbc523-15c2-e029-22a7-a08c5d8c8410
ms.date: 06/08/2017
---


# Options.AutoFormatReplaceHyperlinks Property (Word)

 **True** if e-mail addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are automatically formatted whenever Word AutoFormats a document or range. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatReplaceHyperlinks**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Remarks

Word changes any text that looks like an e-mail address, UNC, or URL to a hyperlink. Word doesn't check the validity of the hyperlink.


## Example

This example enables replacement of any Internet or network paths with hyperlinks, and then it formats the selection automatically.


```vb
Options.AutoFormatReplaceHyperlinks = True 
Selection.Range.AutoFormat
```

This example returns the status of the  **Internet and network paths with hyperlinks** option on the **AutoFormat** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatReplaceHyperlinks
```


## See also


#### Concepts


[Options Object](options-object-word.md)

