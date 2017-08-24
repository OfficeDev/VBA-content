---
title: Rows.Count Property (Publisher)
keywords: vbapb10.chm4915202
f1_keywords:
- vbapb10.chm4915202
ms.prod: publisher
api_name:
- Publisher.Rows.Count
ms.assetid: 790c7616-e9f4-e518-0f4b-6960d144290d
ms.date: 06/08/2017
---


# Rows.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **Rows** object.


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


