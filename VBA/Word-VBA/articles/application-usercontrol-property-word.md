---
title: Application.UserControl Property (Word)
keywords: vbawd10.chm158335077
f1_keywords:
- vbawd10.chm158335077
ms.prod: word
api_name:
- Word.Application.UserControl
ms.assetid: 65cf8ccc-f846-cd86-a8d6-0b2951bad604
ms.date: 06/08/2017
---


# Application.UserControl Property (Word)

 **True** if the document or application was created or opened by the user. Read-only **Boolean** .


## Syntax

 _expression_ . **UserControl**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

The  **UserControl** property returns **False** if the application was created or opened programmatically from another Microsoft Office application with the **Open** method or the **CreateObject** or **GetObject** method.


 **Note**  If Word is visible to the user, or if you call the  **UserControl** property of an **Application** object from within a code module, this property will always return **True** .


## Example

This example displays the status of the  **UserControl** property for the active document. This example will only work correctly when run from another Office application with the Word object library loaded.


```vb
Set wd = New Word.Application 
Set wdDoc = _ 
 wd.Documents.Open("C:\My Documents\doc1.doc") 
If wdDoc.UserControl = True Then 
 MsgBox "This document was created or opened by the user." 
Else 
 MsgBox "This document was created programmatically." 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

