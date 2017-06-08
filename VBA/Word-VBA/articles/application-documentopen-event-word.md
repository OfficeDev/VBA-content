---
title: Application.DocumentOpen Event (Word)
keywords: vbawd10.chm400004
f1_keywords:
- vbawd10.chm400004
ms.prod: word
api_name:
- Word.Application.DocumentOpen
ms.assetid: 21fdd3cd-8769-899e-5749-f64c0e15b265
ms.date: 06/08/2017
---


# Application.DocumentOpen Event (Word)

Occurs when a document is opened.


## Syntax

Private Sub  _expression_ _**DocumentOpen**( **_ByVal Doc As Document_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object declared with events in a class module.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document that's being opened.|

## Remarks

 For more information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


## Example

This example asks the user whether to save all other open documents when a document is opened. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx)for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_DocumentOpen(ByVal Doc As Document) 
 Dim intResponse As Integer 
 Dim strName As String 
 Dim docLoop As Document 
 
 intResponse = MsgBox("Save all other documents?", vbYesNo) 
 
 If intResponse = vbYes Then 
 strName = ActiveDocument.Name 
 For Each docLoop In Documents 
 With docLoop 
 If .Name <> strName Then 
 .Save 
 End If 
 End With 
 Next docLoop 
 End If 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

