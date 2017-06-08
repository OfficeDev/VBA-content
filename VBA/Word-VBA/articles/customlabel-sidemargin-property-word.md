---
title: CustomLabel.SideMargin Property (Word)
keywords: vbawd10.chm152371204
f1_keywords:
- vbawd10.chm152371204
ms.prod: word
api_name:
- Word.CustomLabel.SideMargin
ms.assetid: bd511d0e-36fc-0fd1-57a2-47d9f0a911dc
ms.date: 06/08/2017
---


# CustomLabel.SideMargin Property (Word)

Returns or sets the side margin widths (in points) for the specified custom mailing label. Read/write  **Single** .


## Syntax

 _expression_ . **SideMargin**

 _expression_ An expression that returns a **[CustomLabel](customlabel-object-word.md)** object.


## Remarks

If this property is changed to a value that isn't valid for the specified mailing label layout, an error occurs.


## Example

This example creates a custom label named "VisitorPass" and defines its layout. The left and right margins for each label are 0.75 inch.


```vb
Set myLabel = Application.MailingLabel.CustomLabels _ 
 .Add(Name:="VisitorPass", DotMatrix:=False) 
With myLabel 
 .Height = InchesToPoints(2.17) 
 .HorizontalPitch = InchesToPoints(3.5) 
 .NumberAcross = 2 
 .NumberDown = 4 
 .PageSize = wdCustomLabelLetter 
 .SideMargin = InchesToPoints(0.75) 
 .TopMargin = InchesToPoints(0.17) 
 .VerticalPitch = InchesToPoints(2.17) 
 .Width = InchesToPoints(3.5) 
End With
```


## See also


#### Concepts


[CustomLabel Object](customlabel-object-word.md)

