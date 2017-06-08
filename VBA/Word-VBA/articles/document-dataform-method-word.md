---
title: Document.DataForm Method (Word)
keywords: vbawd10.chm158007402
f1_keywords:
- vbawd10.chm158007402
ms.prod: word
api_name:
- Word.Document.DataForm
ms.assetid: 138f8b31-f076-8573-510f-0295fb612226
ms.date: 06/08/2017
---


# Document.DataForm Method (Word)

Displays the  **Data Form** dialog box, in which you can add, delete, or modify records.


## Syntax

 _expression_ . **DataForm**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

You can use this method with a mail merge main document, a mail merge data source, or any document that contains data delimited by table cells or separator characters.


## Example

This example displays the Data Form dialog box if the active document is a mail merge document.


```vb
If ActiveDocument.MailMerge.State <> wdNormalDocument Then 
 ActiveDocument.DataForm 
End If
```

This example creates a table in a new document and then displays the Data Form dialog box.




```vb
Set aDoc = Documents.Add 
With aDoc 
 .Tables.Add Range:=aDoc.Content, NumRows:=2, NumColumns:=2 
 .Tables(1).Cell(1, 1).Range.Text = "Name" 
 .Tables(1).Cell(1, 2).Range.Text = "Age" 
 .DataForm 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

