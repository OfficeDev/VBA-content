---
title: Application.ActiveDocument Property (Word)
keywords: vbawd10.chm158334979
f1_keywords:
- vbawd10.chm158334979
ms.prod: word
api_name:
- Word.Application.ActiveDocument
ms.assetid: c20a7c9f-f8a4-7913-f53f-10baa6807def
ms.date: 06/08/2017
---


# Application.ActiveDocument Property (Word)

Returns a  **[Document](document-object-word.md)** object that represents the active document (the document with the focus). If there are no documents open, an error occurs. Read-only.


 **Note**  The document in the active protected view window cannot be accessed using this property. Instead, use the [Document](document-object-word.md) property of the **[ActiveProtectedViewWindow](application-activeprotectedviewwindow-property-word.md)** object.


## Syntax

 _expression_ . **ActiveDocument**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example displays the name of the active document, or if there are no documents open, it displays a message.


```vb
If Application.Documents.Count >= 1 Then 
    MsgBox ActiveDocument.Name 
Else 
    MsgBox "No documents are open" 
End If
```

This example collapses the selection to an insertion point and then creates a range for the next five characters in the selection.




```vb
Dim rngTemp As Range 
 
Selection.Collapse Direction:=wdCollapseStart 
Set rngTemp = ActiveDocument.Range(Start:=Selection.Start, _ 
    End:=Selection.Start + 5)
```

This example inserts texts at the beginning of the active document and then prints the document.




```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Range(Start:=0, End:=0) 
With rngTemp 
    .InsertBefore "Company Report" 
    .Font.Name = "Arial" 
    .Font.Size = 24 
    .InsertParagraphAfter 
End With 
 
ActiveDocument.PrintOut
```


## See also


#### Concepts


[Application Object](application-object-word.md)

