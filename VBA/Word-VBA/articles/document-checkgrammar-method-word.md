---
title: Document.CheckGrammar Method (Word)
keywords: vbawd10.chm158007427
f1_keywords:
- vbawd10.chm158007427
ms.prod: word
api_name:
- Word.Document.CheckGrammar
ms.assetid: 980ddb33-94ba-fdae-3c13-6a31fdad3e14
ms.date: 06/08/2017
---


# Document.CheckGrammar Method (Word)

Begins a spelling and grammar check for the specified document or range.


## Syntax

 _expression_ . **CheckGrammar**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

 If the document or range contains errors, this method displays the **Spelling and Grammar** dialog box, with the **Check grammar** check box selected. When applied to a document, this method checks all available stories (such as headers, footers, and text boxes).


## Example

This example begins a spelling and grammar check for all stories in the active document.


```vb
ActiveDocument.CheckGrammar
```


## See also


#### Concepts


[Document Object](document-object-word.md)

