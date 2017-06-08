---
title: Subdocument.HasFile Property (Word)
keywords: vbawd10.chm159973381
f1_keywords:
- vbawd10.chm159973381
ms.prod: word
api_name:
- Word.Subdocument.HasFile
ms.assetid: dbe85127-35cf-7c5f-5ec5-8f1dd35deda1
ms.date: 06/08/2017
---


# Subdocument.HasFile Property (Word)

 **True** if the specified subdocument has been saved to a file. Read-only **Boolean** .


## Syntax

 _expression_ . **HasFile**

 _expression_ A variable that represents a **[Subdocument](subdocument-object-word.md)** object.


## Example

This example displays the file name of each subdocument in the active document. The example also displays a message for each subdocument that has not been saved.


```vb
Dim subLoop As Subdocument 
 
For Each subLoop In ActiveDocument.Subdocuments 
 subLoop.Range.Select 
 If subLoop.HasFile = True Then 
 MsgBox subLoop.Path &; Application.PathSeparator _ 
 &; subLoop.Name 
 Else 
 MsgBox "This subdocument has not been saved." 
 End If 
Next subLoop
```


## See also


#### Concepts


[Subdocument Object](subdocument-object-word.md)

