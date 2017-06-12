---
title: CustomLabel.VerticalPitch Property (Word)
keywords: vbawd10.chm152371207
f1_keywords:
- vbawd10.chm152371207
ms.prod: word
api_name:
- Word.CustomLabel.VerticalPitch
ms.assetid: 5f1107b7-e521-f022-579c-00f14d93d5f6
ms.date: 06/08/2017
---


# CustomLabel.VerticalPitch Property (Word)

Returns or sets the vertical distance between the top of one mailing label and the top of the next mailing label. Read/write  **Single** .


## Syntax

 _expression_ . **VerticalPitch**

 _expression_ An expression that returns a **[CustomLabel](customlabel-object-word.md)** object.


## Remarks

If this property is changed to a value that isn't valid for the specified mailing label layout, an error occurs.


## Example

This example creates a custom label named "VisitorPass" and defines its layout. The distance between the top edge of one label to the top edge of the next label is 2.17 inches.


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

