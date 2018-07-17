---
title: Chart.Walls Property (Word)
keywords: vbawd10.chm79364152
f1_keywords:
- vbawd10.chm79364152
ms.prod: word
api_name:
- Word.Chart.Walls
ms.assetid: f45ae75a-c96c-4441-af81-aedf23787194
ms.date: 06/08/2017
---


# Chart.Walls Property (Word)

Returns the walls of the 3-D chart. Read-only  **[Walls](walls-object-word.md)** .


## Syntax

 _expression_ . **Walls**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Example

The following example sets the color of the wall border of the first chart in the active document to red. You should run the example on a 3-D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Walls.Border. _ 
 ColorIndex = 3 
 End If 
End With 

```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

