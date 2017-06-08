---
title: Application.DocumentChange Event (Word)
keywords: vbawd10.chm400003
f1_keywords:
- vbawd10.chm400003
ms.prod: word
api_name:
- Word.Application.DocumentChange
ms.assetid: 853cbe7e-e4ac-2dba-826a-695d5c137d48
ms.date: 06/08/2017
---


# Application.DocumentChange Event (Word)

Occurs when a new document is created, when an existing document is opened, or when another document is made the active document.


## Syntax

Private Sub  _expression_ _**DocumentChange**()

 _expression_ A variable that represents an **[Application](application-object-word.md)** object declared with events in a class module.


## Remarks

For information about using events with the  **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


## Example

This example asks the user whether to save all the other open documents when the document focus changes. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx)for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_DocumentChange() 
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

