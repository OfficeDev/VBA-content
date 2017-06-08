---
title: Document.ShowSpellingErrors Property (Word)
keywords: vbawd10.chm158007369
f1_keywords:
- vbawd10.chm158007369
ms.prod: word
api_name:
- Word.Document.ShowSpellingErrors
ms.assetid: 75b24653-f694-a5d7-bbb7-3f75f52d9e60
ms.date: 06/08/2017
---


# Document.ShowSpellingErrors Property (Word)

 **True** if Microsoft Word underlines spelling errors in the document. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowSpellingErrors**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

To view spelling errors in a document, you must set the  **[CheckSpellingAsYouType](options-checkspellingasyoutype-property-word.md)** property to **True** .


## Example

This example sets Word to hide the wavy red line that denotes possible spelling errors in the active document.


```vb
ActiveDocument.ShowSpellingErrors = False
```

This example sets Word to show spelling errors in the active document.




```vb
Options.CheckSpellingAsYouType = True 
ActiveDocument.ShowSpellingErrors = True
```

This example returns the current status of the Hide spelling errors in this document checkbox in the Spelling area on the Spelling &; Grammar tab in the Options dialog box.




```
temp = ActiveDocument.ShowSpellingErrors
```


## See also


#### Concepts


[Document Object](document-object-word.md)

