---
title: Axis.DisplayUnitLabel Property (Word)
keywords: vbawd10.chm113049677
f1_keywords:
- vbawd10.chm113049677
ms.prod: word
api_name:
- Word.Axis.DisplayUnitLabel
ms.assetid: fed46896-2968-8332-13b4-8ad0d609169e
ms.date: 06/08/2017
---


# Axis.DisplayUnitLabel Property (Word)

Returns the  **[DisplayUnitLabel](displayunitlabel-object-word.md)** object for the specified axis. Returns **null** if the **[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-word.md)** property is set to **False** . Read-only.


## Syntax

 _expression_ . **DisplayUnitLabel**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Example

The following example sets the label caption to "Millions" for the value axis of the first chart in the active document, and then it turns off automatic font scaling.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue).DisplayUnitLabel 
 .Caption = "Millions" 
 .AutoScaleFont = False 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

