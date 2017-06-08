---
title: Application.NewDocument Property (Word)
keywords: vbawd10.chm158335430
f1_keywords:
- vbawd10.chm158335430
ms.prod: word
api_name:
- Word.Application.NewDocument
ms.assetid: 2f68f98e-1aad-eeac-59c7-4cd5f9d7ad6a
ms.date: 06/08/2017
---


# Application.NewDocument Property (Word)

Returns a  **NewFile** object that represents a document listed on the **New** tab.


## Syntax

 _expression_ . **NewDocument**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example creates a document list item on the New Document task pane in the New From Existing File section.


```vb
Sub CreateNewDocument() 
 Application.NewDocument.Add FileName:="C:\NewFile.doc", _ 
 Section:=msoNewfromExistingFile, DisplayName:="New File", _ 
 Action:=msoCreateNewFile 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

