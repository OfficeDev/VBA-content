---
title: Options.IgnoreInternetAndFileAddresses Property (Word)
keywords: vbawd10.chm162988310
f1_keywords:
- vbawd10.chm162988310
ms.prod: word
api_name:
- Word.Options.IgnoreInternetAndFileAddresses
ms.assetid: 30894aec-958d-b39c-3ef6-a251837f6bbc
ms.date: 06/08/2017
---


# Options.IgnoreInternetAndFileAddresses Property (Word)

 **True** if file name extensions, MS-DOS paths, e-mail addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are ignored while checking spelling. Read/write **Boolean** .


## Syntax

 _expression_ . **IgnoreInternetAndFileAddresses**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore file names and Internet addresses, and then it checks spelling in the active document.


```vb
Options.IgnoreInternetAndFileAddresses = True 
ActiveDocument.CheckSpelling
```

This example returns the current status of the  **Ignore Internet and file addresses** option on the **Spelling &; Grammar** tab in the **Options** dialog box.




```vb
Dim blnTemp As Boolean 
 
blnTemp = Options.IgnoreInternetAndFileAddresses
```


## See also


#### Concepts


[Options Object](options-object-word.md)

