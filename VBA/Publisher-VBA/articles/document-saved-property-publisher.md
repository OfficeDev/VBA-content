---
title: Document.Saved Property (Publisher)
keywords: vbapb10.chm196649
f1_keywords:
- vbapb10.chm196649
ms.prod: publisher
api_name:
- Publisher.Document.Saved
ms.assetid: d1f4357a-103c-2227-d1bd-50706e1f241c
ms.date: 06/08/2017
---


# Document.Saved Property (Publisher)

Returns  **True** if no changes have been made to a publication since it was last saved. Read-only **Boolean**.


## Syntax

 _expression_. **Saved**

 _expression_A variable that represents a  **Document** object.


### Return Value

Boolean


## Remarks

If the  **Saved** property of a modified publication returns **True**, the user won't be prompted to save changes when closing the publication, and all changes made to it since it was last saved will be lost.


## Example

This example saves the active publication if it has been changed since the last time it was saved.


```vb
Sub Saver() 
 
 With Application.ActiveDocument 
 If Not .Saved And .Path <> "" Then .Save 
 End With 
 
End Sub
```


