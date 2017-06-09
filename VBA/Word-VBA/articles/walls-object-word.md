---
title: Walls Object (Word)
keywords: vbawd10.chm384
f1_keywords:
- vbawd10.chm384
ms.prod: word
api_name:
- Word.Walls
ms.assetid: e98c7218-b944-12bb-caf9-daecee4b6c0c
ms.date: 06/08/2017
---


# Walls Object (Word)

Represents the walls of a 3-D chart. 


## Remarks

This object is not a collection. There is no object that represents a single wall; you must return all the walls as a unit.


## Example

Use the  **[Walls](chart-walls-property-word.md)** property to return the **Walls** object. The following example sets the pattern on the walls for the first chart in the active document. If the chart is not a 3-D chart, this example will fail.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Walls.Interior.Pattern = xlGray75 
 End If 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

