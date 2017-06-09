---
title: Fields.Count Property (Publisher)
keywords: vbapb10.chm6029315
f1_keywords:
- vbapb10.chm6029315
ms.prod: publisher
api_name:
- Publisher.Fields.Count
ms.assetid: a8a6b0d4-b029-0b45-6d76-6fb237c31c97
ms.date: 06/08/2017
---


# Fields.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **Fields** object.


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


