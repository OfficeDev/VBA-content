---
title: Language.NameLocal Property (Word)
keywords: vbawd10.chm158138368
f1_keywords:
- vbawd10.chm158138368
ms.prod: word
api_name:
- Word.Language.NameLocal
ms.assetid: b1e91f5e-4ed3-2361-e190-656b0279e8a1
ms.date: 06/08/2017
---


# Language.NameLocal Property (Word)

Returns the name of a proofing tool language in the language of the user. Read-only  **String** .


## Syntax

 _expression_ . **NameLocal**

 _expression_ Required. A variable that represents a **[Language](language-object-word.md)** object.


## Example

This example displays the name of the German language two different ways — first in the language of the user, and then in German.


```vb
MsgBox Languages(wdGerman).NameLocal 
MsgBox Languages(wdGerman).Name
```


## See also


#### Concepts


[Language Object](language-object-word.md)

