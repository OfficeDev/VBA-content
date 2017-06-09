---
title: Document.CodeName Property (Word)
keywords: vbawd10.chm158007558
f1_keywords:
- vbawd10.chm158007558
ms.prod: word
api_name:
- Word.Document.CodeName
ms.assetid: 684f885d-9468-9bc9-d381-ef73286330ff
ms.date: 06/08/2017
---


# Document.CodeName Property (Word)

Returns the code name for the specified document. Read-only  **String** .


## Syntax

 _expression_ . **CodeName**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The code name is the name for the module that houses event macros for a document. The default name for the module is "ThisDocument"; you can view it in the Project window. For information about using events with the Document object, see [Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).


## Example

This example returns the name of the code window for the active document.


```
Msgbox ActiveDocument.CodeName
```


## See also


#### Concepts


[Document Object](document-object-word.md)

