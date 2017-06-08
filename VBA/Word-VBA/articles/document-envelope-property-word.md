---
title: Document.Envelope Property (Word)
keywords: vbawd10.chm158007324
f1_keywords:
- vbawd10.chm158007324
ms.prod: word
api_name:
- Word.Document.Envelope
ms.assetid: 00978466-69b0-a6b8-6111-5b133dd820d5
ms.date: 06/08/2017
---


# Document.Envelope Property (Word)

Returns an  **[Envelope](envelope-object-word.md)** object that represents an envelope and envelope features in a document. Read-only.


## Syntax

 _expression_ . **Envelope**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets the default envelope size to C4 (229 x 324 mm).


```vb
ActiveDocument.Envelope.DefaultSize = "C4"
```

This example displays the delivery address if an envelope has been added to the document; otherwise, a message box is displayed.




```vb
On Error GoTo errhandler 
addr = ActiveDocument.Envelope.Address.Text 
MsgBox Prompt:=addr, Title:="Delivery Address" 
errhandler: 
If Err = 5852 Then MsgBox "Add an envelope to the document"
```

This example creates a new document and adds an envelope with a predefined delivery address and return address.




```
addr = "Don Funk" &; vbCr &; "123 Skye St." _ 
 &; vbCr &; "Our Town, WA 98040" 
retaddr = "Karin Gallagher" &; vbCr &; "123 Main" _ 
 &; vbCr &; "Other Town, WA 98004" 
Documents.Add.Envelope.Insert Address:=addr, ReturnAddress:=retaddr 
ActiveDocument.ActiveWindow.View.Type = wdPrintView
```


## See also


#### Concepts


[Document Object](document-object-word.md)

