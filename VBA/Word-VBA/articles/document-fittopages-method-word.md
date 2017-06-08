---
title: Document.FitToPages Method (Word)
keywords: vbawd10.chm158007400
f1_keywords:
- vbawd10.chm158007400
ms.prod: word
api_name:
- Word.Document.FitToPages
ms.assetid: 8935d286-61b7-432e-ed79-b85708dd1a01
ms.date: 06/08/2017
---


# Document.FitToPages Method (Word)

Decreases the font size of text just enough so that the document will fit on one fewer pages.


## Syntax

 _expression_ . **FitToPages**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

An error occurs if Word is unable to reduce the page count by one.


## Example

This example attempts to reduce the page count of the active document by one page.


```vb
On Error GoTo errhandler 
ActiveDocument.FitToPages 
errhandler: 
If Err = 5538 Then MsgBox "Fit to pages failed"
```

This example attempts to reduce the page count of each open document by one page.




```vb
For Each doc In Documents 
 doc.FitToPages 
Next doc
```


## See also


#### Concepts


[Document Object](document-object-word.md)

