---
title: ShapeNodes.Count Property (Publisher)
keywords: vbapb10.chm3473411
f1_keywords:
- vbapb10.chm3473411
ms.prod: publisher
api_name:
- Publisher.ShapeNodes.Count
ms.assetid: 5b259584-0aad-57bd-4848-cc7f6e96d430
ms.date: 06/08/2017
---


# ShapeNodes.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **ShapeNodes** object.


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


