---
title: CellRange.Count Property (Publisher)
keywords: vbapb10.chm5177347
f1_keywords:
- vbapb10.chm5177347
ms.prod: publisher
api_name:
- Publisher.CellRange.Count
ms.assetid: b21dfbc8-fa1d-aa25-c8a2-ed81629b5da1
ms.date: 06/08/2017
---


# CellRange.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **CellRange** object.


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


