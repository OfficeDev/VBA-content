---
title: Axis.MajorUnit Property (Word)
keywords: vbawd10.chm113049620
f1_keywords:
- vbawd10.chm113049620
ms.prod: word
api_name:
- Word.Axis.MajorUnit
ms.assetid: abfe244f-2718-dc5d-ebc0-d276ee274231
ms.date: 06/08/2017
---


# Axis.MajorUnit Property (Word)

Returns or sets the major units for the value axis. Read/write  **Double** .


## Syntax

 _expression_ . **MajorUnit**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Setting this property sets the  **[MajorUnitIsAuto](axis-majorunitisauto-property-word.md)** property to **False** .

Use the  **[TickMarkSpacing](axis-tickmarkspacing-property-word.md)** property to set tick mark spacing on the category axis.


## Example

The following example sets the major and minor units for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .MajorUnit = 100 
 .MinorUnit = 20 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

