---
title: Pages.Count Property (Publisher)
keywords: vbapb10.chm458755
f1_keywords:
- vbapb10.chm458755
ms.prod: publisher
api_name:
- Publisher.Pages.Count
ms.assetid: 6cc42bb4-4862-6a59-168b-6a97a7e114c8
ms.date: 06/08/2017
---


# Pages.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **Pages** object.


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


