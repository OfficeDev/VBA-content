---
title: Document.SpellingChecked Property (Word)
keywords: vbawd10.chm158007367
f1_keywords:
- vbawd10.chm158007367
ms.prod: word
api_name:
- Word.Document.SpellingChecked
ms.assetid: 053f8fbd-30cd-038f-e36f-d55fdd26fe13
ms.date: 06/08/2017
---


# Document.SpellingChecked Property (Word)

 **True** if spelling has been checked throughout the specified range or document. **False** if all or some of the range or document has not been checked for spelling. Read/write **Boolean** .


## Syntax

 _expression_ . **SpellingChecked**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

To recheck the spelling in a range or document, set the  **SpellingChecked** property to **False** .

To see whether the range or document contains spelling errors, use the  **SpellingErrors** property.


## Example

This example sets the  **SpellingChecked** property to **False** for MyDocument.doc, and then it runs another spelling check on the document.


```vb
Documents("MyDocument.doc").SpellingChecked = False 
Documents("MyDocument.doc").CheckSpelling IgnoreUppercase:=False
```


## See also


#### Concepts


[Document Object](document-object-word.md)

