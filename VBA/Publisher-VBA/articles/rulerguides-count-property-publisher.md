---
title: RulerGuides.Count Property (Publisher)
keywords: vbapb10.chm720899
f1_keywords:
- vbapb10.chm720899
ms.prod: publisher
api_name:
- Publisher.RulerGuides.Count
ms.assetid: 92a93b1a-80c1-7a41-cb94-ac0859a4a470
ms.date: 06/08/2017
---


# RulerGuides.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **RulerGuides** object.


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


