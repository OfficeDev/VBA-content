---
title: WebListBoxItems.Count Property (Publisher)
keywords: vbapb10.chm4128771
f1_keywords:
- vbapb10.chm4128771
ms.prod: publisher
api_name:
- Publisher.WebListBoxItems.Count
ms.assetid: a306e5d1-c0e4-86f3-745a-720f91bf1f25
ms.date: 06/08/2017
---


# WebListBoxItems.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **WebListBoxItems** object.


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


