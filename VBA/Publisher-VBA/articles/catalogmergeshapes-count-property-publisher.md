---
title: CatalogMergeShapes.Count Property (Publisher)
keywords: vbapb10.chm8388611
f1_keywords:
- vbapb10.chm8388611
ms.prod: publisher
api_name:
- Publisher.CatalogMergeShapes.Count
ms.assetid: a871af2f-183c-f5a8-7ad0-c8d25c71e41f
ms.date: 06/08/2017
---


# CatalogMergeShapes.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **CatalogMergeShapes** object.


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


