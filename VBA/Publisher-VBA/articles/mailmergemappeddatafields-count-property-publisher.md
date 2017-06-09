---
title: MailMergeMappedDataFields.Count Property (Publisher)
keywords: vbapb10.chm6488067
f1_keywords:
- vbapb10.chm6488067
ms.prod: publisher
api_name:
- Publisher.MailMergeMappedDataFields.Count
ms.assetid: 45bb34e6-3b6f-2daa-d782-2bbd02b1e7b4
ms.date: 06/08/2017
---


# MailMergeMappedDataFields.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **MailMergeMappedDataFields** object.


## Example

This example displays the number of pages in the active document.


```vb
Sub CountNumberOfPages() 
 MsgBox "Your publication contains " &; _ 
 ActiveDocument.Pages.Count &; " page(s)." 
End Sub
```

This example displays the number of shapes in the active document.




```vb
Sub CountNumberOfShapes() 
 Dim intShapes As Integer 
 Dim pg As Page 
 
 For Each pg In ActiveDocument.Pages 
 intShapes = intShapes + pg.Shapes.Count 
 Next 
 
 MsgBox "Your publication contains " &; intShapes &; " shape(s)." 
End Sub
```


