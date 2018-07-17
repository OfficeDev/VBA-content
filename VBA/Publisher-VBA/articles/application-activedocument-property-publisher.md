---
title: Application.ActiveDocument Property (Publisher)
keywords: vbapb10.chm131073
f1_keywords:
- vbapb10.chm131073
ms.prod: publisher
api_name:
- Publisher.Application.ActiveDocument
ms.assetid: c6293fa6-291c-d8ce-be54-f8a997b95d2e
ms.date: 06/08/2017
---


# Application.ActiveDocument Property (Publisher)

Returns a  **[Document](document-object-publisher.md)** object that represents the active publication. If there are no documents open, an error occurs.


## Syntax

 _expression_. **ActiveDocument**

 _expression_A variable that represents an  **Application** object.


### Return Value

Document


## Example

This example allows the user to assign a file name to the active publication and save it with the new file name. The file name, along with other text, is then inserted after the currently selected text. (Note that Filename must be replaced with a valid publication name for this example to work.)


```vb
Sub NewsLetterSave() 
 
 Dim strFileName As String 
 
 ' Assign the explicit file name to a variable. 
 strFileName = "Filename" 
 Publisher.ActiveDocument.SaveAs strFileName 
 
 ' Insert the file name and supporting text after selected text. 
 Selection.TextRange.Collapse pbCollapseEnd 
 Selection.TextRange = _ 
 " This publication has been saved as " &; strFileName 
 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

