---
title: CoAuthor.Name Property (Word)
keywords: vbawd10.chm81068032
f1_keywords:
- vbawd10.chm81068032
ms.prod: word
api_name:
- Word.CoAuthor.Name
ms.assetid: d9d27cd8-e152-b5a3-286f-3e1b13d09696
ms.date: 06/08/2017
---


# CoAuthor.Name Property (Word)

Returns a  **String** that contains the display name of the specified co author. Read-only.


## Syntax

 _expression_ . **Name**

 _expression_ An expression that returns a **CoAuthor** object.


## Example

The following code example displays the name of the first co author in the active document.


```vb
Set coAuth = ActiveDocument.CoAuthoring.Authors(1) 
MsgBox "The name of the user is " &; _ 
coAuth.Name &; "."
```


## See also


#### Concepts


[CoAuthor Object](coauthor-object-word.md)

