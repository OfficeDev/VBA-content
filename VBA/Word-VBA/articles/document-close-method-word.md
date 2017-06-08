---
title: Document.Close Method (Word)
keywords: vbawd10.chm158008401
f1_keywords:
- vbawd10.chm158008401
ms.prod: word
api_name:
- Word.Document.Close
ms.assetid: 59603a58-17ee-bc65-597b-6200e8be9fbc
ms.date: 06/08/2017
---


# Document.Close Method (Word)

Closes the specified document.


## Syntax

 _expression_ . **Close**( **_SaveChanges_** , **_OriginalFormat_** , **_RouteDocument_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Variant**|Specifies the save action for the document. Can be one of the following  **[WdSaveOptions](wdsaveoptions-enumeration-word.md)** constants: **wdDoNotSaveChanges** , **wdPromptToSaveChanges** , or **wdSaveChanges** .|
| _OriginalFormat_|Optional| **Variant**|Specifies the save format for the document. Can be one of the following  **[WdOriginalFormat](wdoriginalformat-enumeration-word.md)** constants: **wdOriginalDocumentFormat** , **wdPromptUser** , or **wdWordDocument** .|
| _RouteDocument_|Optional| **Variant**| **True** to route the document to the next recipient. If the document does not have a routing slip attached, this argument is ignored.|

## Example

This example prompts the user to save the active document before closing it. If the user clicks Cancel, error 4198 (command failed) is trapped and a message is displayed.


```vb
On Error GoTo errorHandler 
ActiveDocument.Close _ 
 SaveChanges:=wdPromptToSaveChanges, _ 
 OriginalFormat:=wdPromptUser 
errorHandler: 
If Err = 4198 Then MsgBox "Document was not closed"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

