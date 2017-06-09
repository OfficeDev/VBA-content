---
title: CustomLabel.NumberAcross Property (Word)
keywords: vbawd10.chm152371209
f1_keywords:
- vbawd10.chm152371209
ms.prod: word
api_name:
- Word.CustomLabel.NumberAcross
ms.assetid: 3e4d9751-c33b-1780-1e4c-95f9202f4fe0
ms.date: 06/08/2017
---


# CustomLabel.NumberAcross Property (Word)

Returns or sets the number of custom mailing labels across a page. Read/write  **Long** .


## Syntax

 _expression_ . **NumberAcross**

 _expression_ An expression that returns a **[CustomLabel](customlabel-object-word.md)** object.


## Remarks

If this property is changed to a value that isn't valid for the specified mailing label layout, an error occurs.


## Example

This example creates a new custom label named "Dept. Labels" and defines the layout, including the number of labels across the page.


```vb
Set myLabel = Application.MailingLabel.CustomLabels _ 
 .Add(Name:="Dept. Labels", DotMatrix:=False) 
With myLabel 
 .Height = InchesToPoints(0.5) 
 .HorizontalPitch = InchesToPoints(2.06) 
 .NumberAcross = 4 
 .NumberDown = 4 
 .PageSize = wdCustomLabelLetter 
 .SideMargin = InchesToPoints(0.28) 
 .TopMargin = InchesToPoints(0.5) 
 .VerticalPitch = InchesToPoints(2) 
 .Width = InchesToPoints(1.75) 
End With
```


## See also


#### Concepts


[CustomLabel Object](customlabel-object-word.md)

