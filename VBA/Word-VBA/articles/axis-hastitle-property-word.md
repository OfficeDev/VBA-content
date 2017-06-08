---
title: Axis.HasTitle Property (Word)
keywords: vbawd10.chm113049615
f1_keywords:
- vbawd10.chm113049615
ms.prod: word
api_name:
- Word.Axis.HasTitle
ms.assetid: fc221c17-bdaf-a6af-b3dd-58ebd681a955
ms.date: 06/08/2017
---


# Axis.HasTitle Property (Word)

 **True** if the axis or chart has a visible title. Read/write **Boolean** .


## Syntax

 _expression_ . **HasTitle**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

An axis title is represented by an  **[AxisTitle](axistitle-object-word.md)** object.


## Example

The following example adds an axis label to the category axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axis(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

