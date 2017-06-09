---
title: Document.ShowGrammaticalErrors Property (Word)
keywords: vbawd10.chm158007368
f1_keywords:
- vbawd10.chm158007368
ms.prod: word
api_name:
- Word.Document.ShowGrammaticalErrors
ms.assetid: b219a212-232c-0edb-d702-88ed4e097940
ms.date: 06/08/2017
---


# Document.ShowGrammaticalErrors Property (Word)

 **True** if grammatical errors are marked by a wavy green line in the specified document. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowGrammaticalErrors**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

To view grammatical errors in your document, you must set the  **[CheckGrammarAsYouType](options-checkgrammarasyoutype-property-word.md)** property to **True** .


## Example

This example sets Word to check for grammatical errors as you type and to display any errors found in the active document.


```vb
Options.CheckGrammarAsYouType = True 
ActiveDocument.ShowGrammaticalErrors = True
```


## See also


#### Concepts


[Document Object](document-object-word.md)

