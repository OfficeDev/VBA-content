---
title: Global.ActiveDocument Property (Word)
keywords: vbawd10.chm163119107
f1_keywords:
- vbawd10.chm163119107
ms.prod: word
api_name:
- Word.Global.ActiveDocument
ms.assetid: ce25921e-7b90-c122-e054-6be678e4a69b
ms.date: 06/08/2017
---


# Global.ActiveDocument Property (Word)

Returns a  **[Document](document-object-word.md)** object that represents the active document (the document with the focus). Read-only.


## Syntax

 _expression_ . **ActiveDocument**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

If there are no documents open, using this property causes an error. 


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


[Global Object](global-object-word.md)

