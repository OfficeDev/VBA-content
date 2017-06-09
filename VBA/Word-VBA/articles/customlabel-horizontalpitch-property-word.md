---
title: CustomLabel.HorizontalPitch Property (Word)
keywords: vbawd10.chm152371208
f1_keywords:
- vbawd10.chm152371208
ms.prod: word
api_name:
- Word.CustomLabel.HorizontalPitch
ms.assetid: 87d0ba81-3298-ffe2-71d3-eef2301e1484
ms.date: 06/08/2017
---


# CustomLabel.HorizontalPitch Property (Word)

Returns or sets the horizontal distance (in points) between the left edge of one custom mailing label and the left edge of the next mailing label. Read/write  **Single** .


## Syntax

 _expression_ . **HorizontalPitch**

 _expression_ A variable that represents a **[CustomLabel](customlabel-object-word.md)** object.


## Remarks

If this property is changed to a value that isn't valid for the specified mailing label layout, an error occurs.


## Example

This example defines the layout of an existing custom label named "Laser labels." The horizontal distance between the left edge of one label and the left edge of the next label is set to 4.19 inches.


```vb
With Application.MailingLabel.CustomLabels("Laser labels") 
 .Height = InchesToPoints(2) 
 .HorizontalPitch = InchesToPoints(4.19) 
 .NumberAcross = 2 
 .NumberDown = 5 
 .PageSize = wdCustomLabelLetter 
 .SideMargin = InchesToPoints(0.16) 
 .TopMargin = InchesToPoints(0.5) 
 .VerticalPitch = InchesToPoints(2) 
 .Width = InchesToPoints(4) 
End With
```


## See also


#### Concepts


[CustomLabel Object](customlabel-object-word.md)

