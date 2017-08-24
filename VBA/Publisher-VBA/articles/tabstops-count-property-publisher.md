---
title: TabStops.Count Property (Publisher)
keywords: vbapb10.chm5570563
f1_keywords:
- vbapb10.chm5570563
ms.prod: publisher
api_name:
- Publisher.TabStops.Count
ms.assetid: 5ba876e2-b1c0-4de9-6942-02e6688aa169
ms.date: 06/08/2017
---


# TabStops.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **TabStops** object.


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


