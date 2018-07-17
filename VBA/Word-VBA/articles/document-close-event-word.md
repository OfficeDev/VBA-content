---
title: Document.Close Event (Word)
keywords: vbawd10.chm4001006
f1_keywords:
- vbawd10.chm4001006
ms.prod: word
api_name:
- Word.Document.Close
ms.assetid: 7758dbae-b624-d3b0-f42c-1404d40ecc78
ms.date: 06/08/2017
---


# Document.Close Event (Word)

Occurs when a document is closed.


## Syntax

Private Sub  _expression_ _**Close**()

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

If the event procedure is stored in a template, the procedure will run when a new document based on that template is closed and when the template itself is closed (after being opened as a document).



For information about using events with a  **Document** object, see[Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).




## Example

This example makes a backup copy of the document on a file server when the document is closed. (The procedure can be stored in the ThisDocument class module of a document or its attached template.)


```vb
Private Sub Document_Close() 
 ActiveDocument.Save 
 ActiveDocument.SaveAs "\\network\backup\" &; ThisDocument.Name 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

