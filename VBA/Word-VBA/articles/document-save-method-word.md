---
title: Document.Save Method (Word)
keywords: vbawd10.chm158007404
f1_keywords:
- vbawd10.chm158007404
ms.prod: word
api_name:
- Word.Document.Save
ms.assetid: 7e329abc-0530-7016-7712-687de2c780a8
ms.date: 06/08/2017
---


# Document.Save Method (Word)

Saves the specified document.


## Syntax

 _expression_ . **Save**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.
 
**Parameters:**

_NoPrompt_ (Optional)

If `true`, then Word automatically saves all documents.
If `false`, then Word prompts the user to save each document that has changed since it was last saved.

_OriginalFormat_ (Optional)

Specifies the way the documents are saved. Can be one of the WdOriginalFormat constants.

## Remarks

If a document has not been saved before, the  **Save As** dialog box prompts the user for a file name.


## Example

This example saves the active document if it has changed since it was last saved.


```vb
If ActiveDocument.Saved = False Then ActiveDocument.Save
```

This example saves each document in the  **Documents** collection without first prompting the user.




```
Documents.Save NoPrompt:=True, _ 
 OriginalFormat:=wdOriginalDocumentFormat
```


## See also


#### Concepts


[Document Object](document-object-word.md)

