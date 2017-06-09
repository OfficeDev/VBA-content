---
title: Application.DocumentBeforeClose Event (Word)
keywords: vbawd10.chm400005
f1_keywords:
- vbawd10.chm400005
ms.prod: word
api_name:
- Word.Application.DocumentBeforeClose
ms.assetid: 91c89b29-3110-85d7-c141-d1add3bb57f1
ms.date: 06/08/2017
---


# Application.DocumentBeforeClose Event (Word)

Occurs immediately before any open document closes.


## Syntax

Private Sub  _expression_ _**DocumentBeforeClose**( **_ByVal Doc As Document_** , **_Cancel As Boolean_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object declared with events in a class module.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **[Document](document-object-word.md)**|The document that's being closed.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the document doesn't close when the procedure is finished.|

## Remarks

 For more information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


## Example

This example prompts the user for a yes or no response before closing any document. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx) for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_DocumentBeforeClose _ 
        (ByVal Doc As Document, _ 
        Cancel As Boolean) 
 
    Dim intResponse As Integer 
 
    intResponse = MsgBox("Do you really " _ 
        &; "want to close the document?", _ 
        vbYesNo) 
 
    If intResponse = vbNo Then Cancel = True 
End Sub
```


