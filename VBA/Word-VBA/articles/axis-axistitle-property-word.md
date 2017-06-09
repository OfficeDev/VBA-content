---
title: Axis.AxisTitle Property (Word)
keywords: vbawd10.chm113049603
f1_keywords:
- vbawd10.chm113049603
ms.prod: word
api_name:
- Word.Axis.AxisTitle
ms.assetid: 6184ae08-780c-0d39-761e-e1b8a4e140cb
ms.date: 06/08/2017
---


# Axis.AxisTitle Property (Word)

Returns the title of the specified axis. Read-only  **[AxisTitle](axistitle-object-word.md)** .


## Syntax

 _expression_ . **AxisTitle**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Example

The following example adds an axis label to the category axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

