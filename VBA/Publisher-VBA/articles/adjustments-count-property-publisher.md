---
title: Adjustments.Count Property (Publisher)
keywords: vbapb10.chm2424835
f1_keywords:
- vbapb10.chm2424835
ms.prod: publisher
api_name:
- Publisher.Adjustments.Count
ms.assetid: 1b32f1c3-0bbc-a175-4f59-36cc76df12fd
ms.date: 06/08/2017
---


# Adjustments.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents an  **Adjustments** object.


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


## See also


#### Concepts


 [Adjustments Object](adjustments-object-publisher.md)

