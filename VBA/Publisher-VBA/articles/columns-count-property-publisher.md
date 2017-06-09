---
title: Columns.Count Property (Publisher)
keywords: vbapb10.chm5046274
f1_keywords:
- vbapb10.chm5046274
ms.prod: publisher
api_name:
- Publisher.Columns.Count
ms.assetid: 2f7fdb6a-6cd0-2ede-bd34-6954ef23c1a0
ms.date: 06/08/2017
---


# Columns.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **Columns** object.


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


