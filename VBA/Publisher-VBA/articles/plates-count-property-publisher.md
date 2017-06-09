---
title: Plates.Count Property (Publisher)
keywords: vbapb10.chm2818050
f1_keywords:
- vbapb10.chm2818050
ms.prod: publisher
api_name:
- Publisher.Plates.Count
ms.assetid: f042ff71-c649-e4a9-eb69-9d2b084b6e56
ms.date: 06/08/2017
---


# Plates.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **Plates** object.


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


