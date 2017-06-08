---
title: Language.ID Property (Word)
keywords: vbawd10.chm158138378
f1_keywords:
- vbawd10.chm158138378
ms.prod: word
api_name:
- Word.Language.ID
ms.assetid: 8af15ba5-19f0-2a65-e44a-a9fed55f8239
ms.date: 06/08/2017
---


# Language.ID Property (Word)

Returns a number that identifies the specified language. Read-only  **WdLanguageID** .


## Syntax

 _expression_ . **ID**

 _expression_ Required. A variable that represents a **[Language](language-object-word.md)** object.


## Example

This example formats the selection with the Icelandic proofing tools language.


```
Selection.LanguageID = Languages("Icelandic").ID
```


## See also


#### Concepts


[Language Object](language-object-word.md)

