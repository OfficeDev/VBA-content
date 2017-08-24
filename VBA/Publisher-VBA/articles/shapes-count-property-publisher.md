---
title: Shapes.Count Property (Publisher)
keywords: vbapb10.chm2162691
f1_keywords:
- vbapb10.chm2162691
ms.prod: publisher
api_name:
- Publisher.Shapes.Count
ms.assetid: 43052c93-461c-ca6a-3c8c-7142bd6d9ea1
ms.date: 06/08/2017
---


# Shapes.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **Shapes** object.


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


