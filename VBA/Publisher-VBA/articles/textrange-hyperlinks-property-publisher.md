---
title: TextRange.Hyperlinks Property (Publisher)
keywords: vbapb10.chm5308485
f1_keywords:
- vbapb10.chm5308485
ms.prod: publisher
api_name:
- Publisher.TextRange.Hyperlinks
ms.assetid: 0cf1f043-532c-3ffc-67cf-389adc5ac02f
ms.date: 06/08/2017
---


# TextRange.Hyperlinks Property (Publisher)

Returns a  **[Hyperlinks](hyperlinks-object-publisher.md)** collection representing all the hyperlinks in the specified text range.


## Syntax

 _expression_. **Hyperlinks**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

Hyperlinks


## Example

The following example looks for all the shapes on page one of the active publication that have text frames and reports how many hyperlinks each shape has.


```vb
Dim hypAll As Hyperlinks 
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.HasTextFrame = msoTrue Then 
 Set hypAll = shpLoop.TextFrame.TextRange.Hyperlinks 
 Debug.Print "Shape " &; shpLoop.Name _ 
 &; " has " &; hypAll.Count &; " hyperlinks." 
 End If 
Next shpLoop
```


