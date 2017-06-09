---
title: Application.NewDocument Event (Word)
keywords: vbawd10.chm400008
f1_keywords:
- vbawd10.chm400008
ms.prod: word
api_name:
- Word.Application.NewDocument
ms.assetid: afe5b924-3067-69e7-4331-a9ea2b30b9b5
ms.date: 06/08/2017
---


# Application.NewDocument Event (Word)

Occurs when a new document is created.


## Syntax

Private Sub Application **_NewDocument**(ByVal  **_Doc_** As Document)

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The new document.|

## Remarks

For more information about using events with the  **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


## Example

This example asks the user whether to save all other open documents when a new document is created. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx)for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_NewDocument(ByVal Doc As Document) 
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

