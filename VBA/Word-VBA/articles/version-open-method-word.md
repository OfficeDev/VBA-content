---
title: Version.Open Method (Word)
keywords: vbawd10.chm162791527
f1_keywords:
- vbawd10.chm162791527
ms.prod: word
api_name:
- Word.Version.Open
ms.assetid: 97880749-0cf1-21bb-e268-8907e424127a
ms.date: 06/08/2017
---


# Version.Open Method (Word)

Opens the specified version of a document. Returns a  **Document** object representing the opened document.


## Syntax

 _expression_ . **Open**

 _expression_ Required. A variable that represents a **[Version](version-object-word.md)** object.


### Return Value

Document


## Example

This example opens the most recent version of Report.doc.


```vb
Sub OpenVersion() 
 Dim mydoc As Document 
 Set mydoc = Documents.Open("C:\MyFiles\Report.doc") 
 If mydoc.Versions.Count > 0 Then 
 mydoc.Versions(mydoc.Versions.Count).Open 
 Else 
 MsgBox "There are no saved versions for this document." 
 End If 
End Sub
```


## See also


#### Concepts


[Version Object](version-object-word.md)

