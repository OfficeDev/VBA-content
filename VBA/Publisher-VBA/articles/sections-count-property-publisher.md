---
title: Sections.Count Property (Publisher)
keywords: vbapb10.chm7340034
f1_keywords:
- vbapb10.chm7340034
ms.prod: publisher
api_name:
- Publisher.Sections.Count
ms.assetid: 39a8848b-e528-7635-8f02-57f200f6a4c9
ms.date: 06/08/2017
---


# Sections.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **Sections** object.


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


